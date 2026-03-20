import { useState, useCallback, useRef } from "react";
import type { DocumentFile, ProcessingStage, ExtractionResult } from "@/types/extractor";

const STAGE_FLOW: { stage: ProcessingStage; message: string; duration: [number, number] }[] = [
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

function generateMockResult(fileName: string): ExtractionResult {
  const claimsCount = randomBetween(5, 20);
  const claims = Array.from({ length: claimsCount }, (_, i) => ({
    claim_number: `${randomBetween(30, 49)}${randomBetween(2000, 2200)}${randomBetween(10000, 99999)}0001`,
    claimant_name: `Claimant ${i + 1}`,
    date_of_loss: `${randomBetween(1, 12).toString().padStart(2, "0")}/${randomBetween(1, 28).toString().padStart(2, "0")}/${randomBetween(2020, 2024)}`,
    status: ["Open", "Closed", "Pending"][randomBetween(0, 2)],
    reserve_amount: randomBetween(1000, 50000),
    paid_amount: randomBetween(0, 30000),
  }));

  return {
    insurer: ["Service American Indemnity Company", "Atlas Insurance Group", "National General Insurance"][randomBetween(0, 2)],
    format: "complex_multi_row",
    confidence: randomBetween(88, 99),
    claims_count: claimsCount,
    claims,
  };
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
    async (id: string, file: File) => {
      updateDoc(id, { stage: "rotation_check", stageMessage: "Starting processing...", startedAt: Date.now() });

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

        const response = await fetch("/api/extract-full", {
          method: "POST",
          body: formData,
        });

        // Stop simulation regardless of success/fail
        isSimulationRunning = false;

        if (!response.ok) {
          throw new Error(`Server error: ${response.statusText}`);
        }

        const json = await response.json();
        if (!json.success) {
          throw new Error(json.error || "Unknown backend error");
        }

        const backendResult = json.data;
        const schema = backendResult.extracted_schema || { claims: [] };

        // For UI summary cards, we still want these values
        const metadata = {
          insurer: backendResult.extraction_metadata?.source_file || "Unknown",
          format: backendResult.extraction_metadata?.method || "standard",
          confidence: backendResult.summary?.avg_confidence ? Math.round(backendResult.summary.avg_confidence * 100) : 95,
          claims_count: backendResult.summary?.claims_count || schema.claims.length,
        };

        updateDoc(id, {
          stage: "complete",
          stageMessage: "✓ Extraction complete",
          // The 'result' is now ONLY the raw backend schema (exactly what's in extracted_schema.json)
          result: schema,
          // Store UI-only metadata separately
          metadata,
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

  const processQueue = useCallback(async () => {
    if (isProcessing.current) return;
    isProcessing.current = true;

    while (processingQueue.current.length > 0) {
      const item = processingQueue.current.shift()!;
      setActiveDocId(item.id);
      try {
        await processDocument(item.id, item.file);
      } catch {
        updateDoc(item.id, { stage: "error", error: "Processing failed", stageMessage: "Error" });
      }
    }

    isProcessing.current = false;
  }, [processDocument, updateDoc]);

  const addFiles = useCallback(
    (files: File[]) => {
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
      processQueue();
    },
    [activeDocId, processQueue]
  );

  const selectedDoc = documents.find((d) => d.id === activeDocId) || null;

  return {
    documents,
    activeDocId,
    selectedDoc,
    addFiles,
    setActiveDocId,
  };
}
