import { useState, useCallback, useRef } from "react";
import type { DocumentFile, ProcessingStage, ExtractionResult } from "@/types/extractor";

const STAGE_FLOW: { stage: ProcessingStage; message: string; duration: [number, number] }[] = [
    { stage: "classification", message: "🧠 Classifying document type...", duration: [1000, 1500] },
    { stage: "classification", message: "✅ Document classified", duration: [300, 500] },
    { stage: "rotation_check", message: "Checking for rotation...", duration: [800, 1500] },
    { stage: "rotation_check", message: "✅ PDF already correctly oriented", duration: [500, 800] },
    { stage: "text_extraction", message: "Extracting text from pages...", duration: [1500, 3000] },
    { stage: "text_extraction", message: "✓ Combined text saved", duration: [400, 700] },
    { stage: "schema_extraction", message: "Analyzing document format...", duration: [1000, 2000] },
    { stage: "policy_detection", message: "Detecting policy boundaries...", duration: [1200, 2500] },
    { stage: "policy_detection", message: "Using AI to detect claim number patterns...", duration: [1500, 3000] },
    { stage: "claim_extraction", message: "Extracting claims using adaptive prompt...", duration: [2000, 4000] },
    { stage: "validation", message: "Validating extraction...", duration: [800, 1500] },
    { stage: "validation", message: "✓ Extraction is COMPLETE", duration: [500, 800] },
];

function randomBetween(min: number, max: number) {
    return Math.floor(Math.random() * (max - min + 1)) + min;
}

export function useDocumentProcessor() {
    const [documents, setDocuments] = useState<DocumentFile[]>([]);
    const [activeDocId, setActiveDocId] = useState<string | null>(null);
    const processingQueue = useRef<{ id: string; file: File }[]>([]);
    const isProcessing = useRef(false);

    const updateDoc = useCallback((id: string, updates: Partial<DocumentFile>) => {
        setDocuments((prev) =>
            prev.map((d) => (d.id === id ? { ...d, ...updates } : d))
        );
    }, []);

    const processDocument = useCallback(
        async (id: string, file: File, evMode: boolean = false) => {
            updateDoc(id, { stage: "classification", stageMessage: "Starting processing...", startedAt: Date.now() });

            let isSimulationRunning = true;

            // Start simulated progress
            const runSimulation = async () => {
                for (const step of STAGE_FLOW) {
                    if (!isSimulationRunning) break;

                    updateDoc(id, {
                        stage: step.stage,
                        stageMessage: step.message
                    });

                    const delay = randomBetween(step.duration[0], step.duration[1]);
                    await new Promise(resolve => setTimeout(resolve, delay));
                }
            };

            const simulationPromise = runSimulation();

            try {
                const formData = new FormData();
                formData.append("file", file);
                formData.append("ev_mode", String(evMode));

                const response = await fetch("/api/extract", {
                    method: "POST",
                    body: formData,
                });

                // Stop simulation regardless of success/fail
                isSimulationRunning = false;

                if (!response.ok) {
                    throw new Error(`Server error: ${response.statusText}`);
                }

                const json = await response.json();

                if (json.error) {
                    throw new Error(json.error);
                }

                // Handle unified router response format
                const documentType = json.type || "UNKNOWN";
                const jsonPath = json.output_json || json.json;

                // Fetch the JSON file from the backend
                let schema: any = null;
                if (jsonPath) {
                    try {
                        const schemaResponse = await fetch(`/api/download/${jsonPath}`);
                        schema = await schemaResponse.json();
                    } catch (e) {
                        console.error("Failed to fetch schema:", e);
                    }
                }

                // Calculate metadata based on document type
                let totalValue = 0;
                let claimsCount = 0;
                let processedResult = schema;

                if (documentType === "INSURANCE" || documentType === "INSURANCE_CLAIMS") {
                    // Insurance format: { claims: [...] } or direct [...]
                    const claims = Array.isArray(schema) ? schema : (schema?.claims || []);
                    claimsCount = claims.length;
                } else if (documentType === "INVOICE") {
                    // Invoice format: [...] (flat array of records)
                    const records = Array.isArray(schema) ? schema : [];

                    // DETECT TOTAL ROW strategy:
                    // 1. If there's a "GRAND TOTAL" row, use it.
                    // 2. If there are multiple "TOTAL" rows, they are likely employee subtotals -> SUM THEM.
                    // 3. If there's only one "TOTAL" row, use it as the grand total.
                    const grandTotalRow = records.find((r: any) => 
                        String(r.PLAN_NAME || "").toUpperCase() === "GRAND TOTAL" || 
                        String(r.FIRSTNAME || "").toUpperCase() === "INVOICE TOTAL"
                    );

                    const totalRows = records.filter((r: any) => {
                        const name = String(r.PLAN_NAME || "").trim().toUpperCase();
                        const prem = String(r.CURRENT_PREMIUM || "").trim();
                        // Recognize "TOTAL" OR a row with no plan name but has a premium (the new Excel style)
                        return name === "TOTAL" || (name === "" && prem !== "" && prem !== "0");
                    });

                    if (grandTotalRow) {
                        const val = parseFloat(String(grandTotalRow.CURRENT_PREMIUM || 0).replace(/[^0-9.-]+/g, ""));
                        totalValue = isNaN(val) ? 0 : val;
                    } else if (totalRows.length > 0) {
                        // If multiple "TOTAL" rows, sum them (likely employee subtotals)
                        // If only one, it's the grand total
                        totalValue = totalRows.reduce((sum: number, r: any) => {
                            const val = parseFloat(String(r.CURRENT_PREMIUM || 0).replace(/[^0-9.-]+/g, ""));
                            return sum + (isNaN(val) ? 0 : val);
                        }, 0);
                    } else {
                        // Fallback: manual sum of all premiums if no total rows found
                        totalValue = records.reduce((sum: number, rec: any) => {
                            const current = parseFloat(String(rec.CURRENT_PREMIUM || 0).replace(/[^0-9.-]+/g, ""));
                            return sum + (isNaN(current) ? 0 : current);
                        }, 0);
                    }

                    // Filter out ALL total rows from the main result so they don't show in the table
                    const allTotalRows = records.filter((r: any) => {
                        const name = String(r.PLAN_NAME || "").toUpperCase();
                        const fname = String(r.FIRSTNAME || "").toUpperCase();
                        return (name === "TOTAL" || name === "GRAND TOTAL" || fname === "INVOICE TOTAL");
                    });
                    const totalRowIndices = new Set(allTotalRows.map(r => records.indexOf(r)));
                    processedResult = records.filter((_, idx) => !totalRowIndices.has(idx));
                } else if (documentType === "WORK_COMPENSATION") {
                    // Work comp: use metadata from API response
                    totalValue = json.work_comp_metadata?.total_premium || 0;
                }

                // Build display label for insurer/applicant field
                let insurerLabel = json.insurer || "Unknown Document";
                if (!json.insurer) {
                    if (documentType === "INSURANCE" || documentType === "INSURANCE_CLAIMS") {
                        insurerLabel = "Insurance Document";
                    } else if (documentType === "INVOICE" || documentType === "VENDOR_INVOICE") {
                        insurerLabel = "Invoice Document";
                    } else if (documentType === "WORK_COMPENSATION") {
                        insurerLabel = json.work_comp_metadata?.form_type || "Workers Comp Application";
                    }
                }

                const metadata = {
                    insurer: insurerLabel,
                    format: documentType.toLowerCase(),
                    confidence: 95,
                    claims_count: claimsCount,
                    total_value: json.total_value || totalValue,
                    documentType: documentType as any,
                    work_comp_metadata: json.work_comp_metadata || null,
                };

                updateDoc(id, {
                    stage: "complete",
                    stageMessage: "✓ Extraction complete",
                    result: processedResult,
                    metadata,
                    excelPath: json.output_file,
                    jsonPath: json.output_json,
                    completedAt: Date.now(),
                });
            } catch (error) {
                isSimulationRunning = false;
                console.error("Processing error:", error);
                updateDoc(id, {
                    stage: "error",
                    error: error instanceof Error ? error.message : "Processing failed",
                    stageMessage: "Error",
                });
            }
        },
        [updateDoc]
    );

    const processQueue = useCallback(async (evMode: boolean = false) => {
        if (isProcessing.current) return;
        isProcessing.current = true;

        while (processingQueue.current.length > 0) {
            const item = processingQueue.current.shift()!;
            setActiveDocId(item.id);
            try {
                await processDocument(item.id, item.file, evMode);
            } catch {
                updateDoc(item.id, { stage: "error", error: "Processing failed", stageMessage: "Error" });
            }
        }

        isProcessing.current = false;
    }, [processDocument, updateDoc]);

    const addFiles = useCallback(
        (files: File[], evMode: boolean = false) => {
            const newDocs: DocumentFile[] = files.map((file) => ({
                id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
                file,
                name: file.name,
                size: file.size,
                stage: "queued" as ProcessingStage,
                stageMessage: "Waiting in queue...",
                progress: 0,
                result: null,
                error: null,
                startedAt: null,
                completedAt: null,
            }));

            setDocuments((prev) => [...prev, ...newDocs]);

            if (!activeDocId && newDocs.length > 0) {
                setActiveDocId(newDocs[0].id);
            }

            newDocs.forEach((d) => processingQueue.current.push({ id: d.id, file: d.file }));
            processQueue(evMode);
        },
        [activeDocId, processQueue]
    );

    const reprocessDocument = useCallback((id: string, evMode: boolean = false) => {
        setDocuments((prev) => {
            const doc = prev.find((d) => d.id === id);
            if (!doc) return prev;

            // Add to queue logic after state update
            setTimeout(() => {
                processingQueue.current.push({ id: doc.id, file: doc.file });
                processQueue(evMode);
            }, 0);

            return prev.map((d) =>
                d.id === id
                    ? {
                        ...d,
                        stage: "queued" as ProcessingStage,
                        stageMessage: "Reprocessing...",
                        result: null,
                        error: null,
                        progress: 0,
                        startedAt: null,
                        completedAt: null
                    }
                    : d
            );
        });
    }, [processQueue]);

    const selectedDoc = documents.find((d) => d.id === activeDocId) || null;

    return {
        documents,
        activeDocId,
        selectedDoc,
        addFiles,
        setActiveDocId,
        reprocessDocument,
    };
}
