import { useState, useCallback } from "react";
import AppHeader from "@/components/AppHeader";
import UploadArea from "@/components/UploadArea";
import DocumentQueue from "@/components/DocumentQueue";
import ResultsPanel from "@/components/ResultsPanel";
import MergeJsonButton from "@/components/MergeJsonButton";
import ClaimSummary from "@/components/ClaimSummary";
import ProcessingStages from "@/components/ProcessingStages";
import { Switch } from "@/components/ui/switch";
import { Label } from "@/components/ui/label";
import { useDocumentProcessor } from "@/hooks/useDocumentProcessor";
import { FileText } from "lucide-react";
import { toast } from "sonner";

const Index = () => {
  const { documents, activeDocId, selectedDoc, addFiles, setActiveDocId, reprocessDocument } =
    useDocumentProcessor();
  const [mergedSummary, setMergedSummary] = useState<string | null>(null);
  const [isMerging, setIsMerging] = useState(false);
  const [evMode, setEvMode] = useState(false);

  const getMergeCategory = (type?: string) => {
    if (!type) return "UNKNOWN";
    if (type === "INSURANCE" || type === "INSURANCE_CLAIMS") return "INSURANCE";
    if (type === "WORK_COMPENSATION") return "WORK_COMPENSATION";
    if (type === "INVOICE") return "INVOICE";
    return type;
  };

  const completedDocs = documents.filter((d) => d.stage === "complete" && d.result);

  // Smarter merge logic: check if there are multiple docs of the SAME category as the selected one
  const selectedCategory = getMergeCategory(selectedDoc?.metadata?.documentType);
  const docsInSameCategory = completedDocs.filter(d => getMergeCategory(d.metadata?.documentType) === selectedCategory);

  const isMergeable = docsInSameCategory.length >= 2;
  const hasMultipleDocs = completedDocs.length >= 2;
  const sharedType = isMergeable ? selectedCategory : null;

  const getAllClaims = useCallback(() => {
    return docsInSameCategory.flatMap((d) => {
      const result = d.result;
      const metadata = d.metadata;
      const category = getMergeCategory(metadata?.documentType);

      if (category === "WORK_COMPENSATION") {
        const wcData = (result as any)?.data || {};
        return wcData.ratingByState || [];
      }

      if (category === "INSURANCE") {
        return (result as any)?.claims || [];
      }
      return Array.isArray(result) ? result : [];
    });
  }, [docsInSameCategory]);

  const handleDownloadMergedJson = useCallback(() => {
    const claims = getAllClaims();
    const blob = new Blob([JSON.stringify(claims, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `merged_claims_${Date.now()}.json`;
    a.click();
    URL.revokeObjectURL(url);
  }, [getAllClaims]);

  const handleDownloadMergedCsv = useCallback(() => {
    const claims = getAllClaims();
    if (claims.length === 0) return;

    const headers = Array.from(new Set(claims.flatMap(c => Object.keys(c))));
    const csvRows = [
      headers.join(','),
      ...claims.map(row =>
        headers.map(header => {
          const val = row[header];
          const escaped = ('' + (val ?? '')).replace(/"/g, '""');
          return `"${escaped}"`;
        }).join(',')
      )
    ];

    const blob = new Blob([csvRows.join('\n')], { type: "text/csv" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `merged_claims_${Date.now()}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  }, [getAllClaims]);

  const handleTriggerMergeAnalysis = useCallback(async () => {
    const claims = getAllClaims();
    if (claims.length === 0) {
      toast.error("No claims data found to analyze");
      return;
    }

    setIsMerging(true);
    try {
      const response = await fetch("/api/claim-summary", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ claims }),
      });

      const data = await response.json();
      if (data.success && data.summary) {
        setMergedSummary(data.summary);
        toast.success("Merged AI Summary generated!");
      } else {
        toast.error("Failed to generate summary: " + (data.error || "Unknown error"));
      }
    } catch (error) {
      console.error("Error generating merged summary:", error);
      toast.error("Error connecting to summary service");
    } finally {
      setIsMerging(false);
    }
  }, [getAllClaims]);

  const hasDocuments = documents.length > 0;

  return (
    <div className="min-h-screen bg-background">
      <AppHeader />

      <main className="max-w-7xl mx-auto px-4 sm:px-8 py-8">
        {/* EV Mode Toggle */}
        <div className="mb-6 flex items-center justify-end gap-3 animated-fade-in">
          <div className="flex flex-col items-end">
            <Label htmlFor="ev-mode" className="text-sm font-bold text-foreground flex items-center gap-2">
              EV (Engage + Velocity Mode)
              {evMode ? (
                <span className="text-[10px] bg-primary/20 text-primary px-1.5 py-0.5 rounded uppercase tracking-wider">Active</span>
              ) : (
                <span className="text-[10px] bg-muted text-muted-foreground px-1.5 py-0.5 rounded uppercase tracking-wider">Inactive</span>
              )}
            </Label>
            <p className="text-[10px] text-muted-foreground mt-0.5 max-w-[200px] text-right">
              {evMode 
                ? "Extracting ONLY Plan Name and Current Premium for ADP/Paychex." 
                : "Extracting complete demographic and coverage fields for Resourcing/Prestige."}
            </p>
          </div>
          <Switch
            id="ev-mode"
            checked={evMode}
            onCheckedChange={setEvMode}
            className="data-[state=checked]:bg-primary"
          />
        </div>

        {/* Upload Section */}
        <section className="mb-8">
          <UploadArea onFilesSelected={(files) => addFiles(files, evMode)} />
        </section>

        {!hasDocuments && (
          <div className="text-center py-16 text-muted-foreground">
            <FileText className="w-16 h-16 mx-auto mb-4 opacity-20" />
            <p className="text-lg font-medium">No documents uploaded yet</p>
            <p className="text-sm mt-1">
              Upload one or more PDF files to begin extraction
            </p>
          </div>
        )}

        {hasDocuments && (
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
            {/* Left sidebar: Document Queue */}
            <aside className="lg:col-span-4 space-y-6">
              <DocumentQueue
                documents={documents}
                activeDocId={activeDocId}
                onSelectDoc={setActiveDocId}
              />

              {/* Active processing stages for selected doc */}
              {selectedDoc &&
                selectedDoc.stage !== "queued" &&
                selectedDoc.stage !== "complete" && (
                  <div className="p-4 rounded-xl border border-border bg-card">
                    <h3 className="text-sm font-bold text-foreground mb-3">
                      Processing: {selectedDoc.name}
                    </h3>
                    <ProcessingStages
                      currentStage={selectedDoc.stage}
                      stageMessage={selectedDoc.stageMessage}
                    />
                  </div>
                )}

              {/* Merge button */}
              <MergeJsonButton
                completedDocsCount={docsInSameCategory.length}
                isMergeable={isMergeable}
                sharedType={sharedType || undefined}
                onDownloadJson={handleDownloadMergedJson}
                onDownloadCsv={handleDownloadMergedCsv}
                onTriggerMergeAnalysis={handleTriggerMergeAnalysis}
                isAnalyzing={isMerging}
              />
            </aside>

            {/* Right panel: Results */}
            <section className="lg:col-span-8">
              <div className="p-6 rounded-xl border border-border bg-card min-h-[400px]">
                <h2 className="text-base font-bold text-foreground mb-4 flex items-center gap-2">
                  <FileText className="w-4 h-4 text-primary" />
                  {selectedDoc
                    ? `Results — ${selectedDoc.name}`
                    : "Extraction Results"}
                </h2>
                {!isMergeable && hasMultipleDocs && (
                  <div className="mb-4 p-2 rounded bg-amber-500/10 border border-amber-500/20 text-[10px] text-amber-600 flex items-center gap-2 italic">
                    <span className="font-bold">Note:</span> Mixed document types (e.g., Invoice + Loss Run) cannot be merged or cross-analyzed.
                  </div>
                )}
                <ResultsPanel
                  document={selectedDoc}
                  onReprocess={(id) => reprocessDocument(id, evMode)}
                  mergedSummary={mergedSummary}
                  isMerging={isMerging}
                  onTriggerMergeAnalysis={handleTriggerMergeAnalysis}
                  onDownloadMergedJson={handleDownloadMergedJson}
                  onDownloadMergedCsv={handleDownloadMergedCsv}
                  hasMultipleDocs={hasMultipleDocs && isMergeable}
                />
              </div>
            </section>
          </div>
        )}
      </main>

      {/* Footer */}
      <footer className="text-center py-6 text-xs text-muted-foreground">
        Data Retrieval Ingestion Verification Engine • AI-Powered PDF Processing
      </footer>
    </div>
  );
};

export default Index;
