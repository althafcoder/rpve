import { useState, useEffect } from "react";
import { Download, FileJson, BarChart3, Building2, RotateCcw, HardHat, FileText, Table as TableIcon, Loader2, Brain, Copy, Merge } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import type { DocumentFile } from "@/types/extractor";
import JsonViewer from "./JsonViewer";
import TableView from "./TableView";
import { toast } from "sonner";

interface ResultsPanelProps {
  document: DocumentFile | null;
  onReprocess?: (id: string) => void;
  mergedSummary?: string | null;
  isMerging?: boolean;
  onTriggerMergeAnalysis?: () => void;
  onDownloadMergedJson?: () => void;
  onDownloadMergedCsv?: () => void;
  hasMultipleDocs?: boolean;
}

const ResultsPanel = ({
  document,
  onReprocess,
  mergedSummary,
  isMerging,
  onTriggerMergeAnalysis,
  onDownloadMergedJson,
  onDownloadMergedCsv,
  hasMultipleDocs
}: ResultsPanelProps) => {
  const [summaryText, setSummaryText] = useState<string | null>(null);
  const [isSummarizing, setIsSummarizing] = useState(false);
  const [activeTab, setActiveTab] = useState("table");

  useEffect(() => {
    setActiveTab("table");
    setSummaryText(null);
  }, [document?.id]);

  if (!document) {
    return (
      <div className="flex flex-col items-center justify-center h-64 text-muted-foreground gap-3">
        <FileJson className="w-12 h-12 opacity-30" />
        <p className="text-sm">Select a document to view results</p>
      </div>
    );
  }

  if (document.stage === "error") {
    return (
      <div className="p-6 bg-destructive/5 rounded-lg border border-destructive/20 flex flex-col gap-4">
        <div>
          <p className="text-sm text-destructive font-medium">Processing Error</p>
          <p className="text-xs text-muted-foreground mt-1">{document.error}</p>
        </div>
        {onReprocess && (
          <Button
            variant="outline"
            size="sm"
            className="w-fit text-destructive border-destructive/20 hover:bg-destructive/10"
            onClick={() => onReprocess(document.id)}
          >
            <RotateCcw className="w-4 h-4 mr-2" />
            Retry Extraction
          </Button>
        )}
      </div>
    );
  }

  if (!document.result) {
    return (
      <div className="flex flex-col items-center justify-center h-64 text-muted-foreground gap-3">
        <div className="stage-pulse">
          <FileJson className="w-12 h-12 opacity-40" />
        </div>
        <p className="text-sm">Processing in progress...</p>
        <p className="text-xs">{document.stageMessage}</p>
      </div>
    );
  }

  const { result, metadata } = document;
  const docType = metadata?.documentType;
  const isWorkComp = docType === "WORK_COMPENSATION";
  const isInvoice = docType === "INVOICE" || docType === "VENDOR_INVOICE";
  const isLossRun = docType === "INSURANCE_CLAIMS";
  const wcMeta = metadata?.work_comp_metadata;

  // Normalized data for table view
  const getTableData = () => {
    if (isLossRun || docType === "INSURANCE") {
      return Array.isArray(result?.claims) ? result.claims : (Array.isArray(result) ? result : []);
    }
    if (isWorkComp) {
      // For Worker Comp, we usually want to show the Rating table as the primary grid
      const wcData = result?.data || {};
      return Array.isArray(wcData.ratingByState) ? wcData.ratingByState : [];
    }
    if (isInvoice) {
      if (docType === "VENDOR_INVOICE") {
        return Array.isArray(result?.LINE_ITEMS) ? result.LINE_ITEMS : [];
      }
      return Array.isArray(result) ? result : [];
    }
    return Array.isArray(result) ? result : [];
  };

  const tableData = getTableData();

  const handleDownloadJson = () => {
    const blob = new Blob([JSON.stringify(result, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = globalThis.document.createElement("a");
    a.href = url;
    a.download = `${document.name.replace(".pdf", "")}_extracted.json`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const handleDownloadExcel = async () => {
    if (!document.excelPath) {
      console.error("No Excel file path available");
      return;
    }

    try {
      const response = await fetch(`/api/download/${document.excelPath}`);
      if (!response.ok) throw new Error("Download failed");

      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      const a = globalThis.document.createElement("a");
      a.href = url;
      a.download = document.excelPath;
      a.click();
      URL.revokeObjectURL(url);
    } catch (error) {
      console.error("Excel download error:", error);
    }
  };

  const handleAnalyze = async () => {
    const claims = Array.isArray(result?.claims) ? result.claims : Array.isArray(result) ? result : [];
    if (claims.length === 0) {
      toast.error("No claims data found to analyze");
      return;
    }

    setIsSummarizing(true);
    setSummaryText(null);
    try {
      const response = await fetch("/api/claim-summary", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ claims }),
      });

      const data = await response.json();
      if (data.success && data.summary) {
        setSummaryText(data.summary);
        toast.success("AI Summary generated!");
      } else {
        toast.error("Failed to generate summary: " + (data.error || "Unknown error"));
      }
    } catch (error) {
      console.error("Error generating summary:", error);
      toast.error("Error connecting to summary service");
    } finally {
      setIsSummarizing(false);
    }
  };

  const handleCopySummary = () => {
    if (summaryText) {
      navigator.clipboard.writeText(summaryText);
      toast.success("Summary copied to clipboard");
    }
  };

  const parseBold = (text: string) => {
    const parts = text.split(/(\*\*.*?\*\*)/g);
    return parts.map((part, i) => {
      if (part.startsWith("**") && part.endsWith("**")) {
        return <strong key={i} className="font-bold text-foreground">{part.slice(2, -2)}</strong>;
      }
      return part;
    });
  };

  const renderMarkdown = (text: string) => {
    const lines = text.split("\n");
    return lines.map((line, index) => {
      const trimmedLine = line.trim();
      if (!trimmedLine) return <div key={index} className="h-2" />;
      if (trimmedLine.startsWith("### ")) return <h3 key={index} className="text-sm font-bold mt-4 mb-2 text-primary">{trimmedLine.slice(4)}</h3>;
      if (trimmedLine.startsWith("## ")) return <h2 key={index} className="text-base font-bold mt-5 mb-3 text-primary border-b border-border pb-1">{trimmedLine.slice(3)}</h2>;
      if (trimmedLine.startsWith("# ")) return <h1 key={index} className="text-lg font-bold mt-6 mb-4 text-primary">{trimmedLine.slice(2)}</h1>;
      if (trimmedLine.startsWith("- ") || trimmedLine.startsWith("* ")) {
        return (
          <div key={index} className="flex gap-2 ml-2 my-1">
            <span className="text-primary mt-1.5">•</span>
            <span>{parseBold(trimmedLine.slice(2))}</span>
          </div>
        );
      }
      return <p key={index} className="my-1">{parseBold(line)}</p>;
    });
  };

  // Dynamic tab configuration
  const showSummary = isLossRun;
  const showMerge = hasMultipleDocs && isLossRun;
  const tabCount = 2 + (showSummary ? 1 : 0) + (showMerge ? 1 : 0);

  return (
    <div className="space-y-4 animate-slide-up">
      {/* Summary cards */}
      {isWorkComp ? (
        /* Work Comp: only Form Type + Confidence */
        <div className="grid grid-cols-2 gap-3">
            <div className="p-3 rounded-lg bg-muted/50 border border-border">
            <div className="flex items-center gap-2 mb-1">
              <HardHat className="w-3.5 h-3.5 text-primary" />
              <span className="text-[11px] text-muted-foreground font-medium">Form Type</span>
            </div>
            <p className="text-sm font-semibold text-foreground truncate">
              {wcMeta?.form_type || metadata?.insurer || "N/A"}
            </p>
          </div>
          <div className="p-3 rounded-lg bg-muted/50 border border-border">
            <div className="flex items-center gap-2 mb-1">
              <FileText className="w-3.5 h-3.5 text-primary" />
              <span className="text-[11px] text-muted-foreground font-medium">Confidence</span>
            </div>
            <p className="text-sm font-semibold text-foreground">
              {metadata?.confidence ? `${metadata.confidence}%` : "N/A"}
            </p>
          </div>
        </div>
      ) : (
        /* Insurance / Invoice: original 3-card layout */
        <div className="grid grid-cols-3 gap-3">
          <div className="p-3 rounded-lg bg-muted/50 border border-border">
            <div className="flex items-center gap-2 mb-1">
              <Building2 className="w-3.5 h-3.5 text-primary" />
              <span className="text-[11px] text-muted-foreground font-medium">Insurer</span>
            </div>
            <p className="text-sm font-semibold text-foreground truncate">{metadata?.insurer || "N/A"}</p>
          </div>
          <div className="p-3 rounded-lg bg-muted/50 border border-border">
            <div className="flex items-center gap-2 mb-1">
              <BarChart3 className="w-3.5 h-3.5 text-primary" />
              <span className="text-[11px] text-muted-foreground font-medium">
                {isInvoice ? "Total Value" : "Claims Found"}
              </span>
            </div>
            <p className="text-sm font-semibold text-foreground">
              {isInvoice
                ? new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' }).format(metadata?.total_value || 0)
                : (() => {
                    if (typeof metadata?.claims_count === "number") {
                      return metadata.claims_count;
                    }
                    const raw: any = result ?? [];
                    const arr: any[] = Array.isArray(raw)
                      ? raw
                      : Array.isArray(raw.claims)
                        ? raw.claims
                        : [];
                    return arr.filter((item) => item && item.claim_number).length;
                  })()
              }
            </p>
          </div>
          <div className="p-3 rounded-lg bg-muted/50 border border-border">
            <div className="flex items-center gap-2 mb-1">
              <FileText className="w-3.5 h-3.5 text-primary" />
              <span className="text-[11px] text-muted-foreground font-medium">Confidence</span>
            </div>
            <p className="text-sm font-semibold text-foreground">
              {metadata?.confidence ? `${metadata.confidence}%` : "N/A"}
            </p>
          </div>
        </div>
      )}

      {/* WC States badge row for Work Comp */}
      {isWorkComp && wcMeta?.wc_states && wcMeta.wc_states.length > 0 && (
        <div className="flex flex-wrap gap-1.5">
          <span className="text-[11px] text-muted-foreground font-medium self-center">States:</span>
          {wcMeta.wc_states.map((state) => (
            <span
              key={state}
              className="px-2 py-0.5 rounded-full text-[11px] font-semibold bg-primary/10 text-primary border border-primary/20"
            >
              {state}
            </span>
          ))}
        </div>
      )}

      {/* Actions */}
      <div className="flex flex-wrap gap-2">
        <Button size="sm" onClick={handleDownloadJson} className="h-8 text-[11px] font-bold">
          <FileJson className="w-3.5 h-3.5 mr-1.5" />
          Download JSON
        </Button>
        <Button size="sm" variant="outline" onClick={handleDownloadExcel} disabled={!document.excelPath} className="h-8 text-[11px] font-bold">
          <Download className="w-3.5 h-3.5 mr-1.5" />
          Download Excel
        </Button>
        {isLossRun && (
          <Button
            size="sm"
            variant="secondary"
            className="h-8 text-[11px] font-bold bg-primary/10 text-primary hover:bg-primary/20 border-primary/20 border"
            onClick={() => {
              setActiveTab("summary");
              if (!summaryText) handleAnalyze();
            }}
            disabled={isSummarizing}
          >
            {isSummarizing ? (
              <Loader2 className="w-3.5 h-3.5 mr-1.5 animate-spin" />
            ) : (
              <Brain className="w-3.5 h-3.5 mr-1.5" />
            )}
            AI Summary
          </Button>
        )}
        {showMerge && (
          <Button
            size="sm"
            variant="secondary"
            className="h-8 text-[11px] font-bold bg-stage-done/10 text-stage-done hover:bg-stage-done/20 border-stage-done/20 border"
            onClick={() => setActiveTab("merge")}
          >
            <Merge className="w-3.5 h-3.5 mr-1.5" />
            Merge Actions
          </Button>
        )}
        {onReprocess && (
          <Button size="sm" variant="ghost" className="h-8 text-[11px] font-bold text-muted-foreground hover:text-primary ml-auto" onClick={() => onReprocess(document.id)}>
            <RotateCcw className="w-3 h-3 mr-1.5" />
            Reprocess
          </Button>
        )}
      </div>

      {/* View Switcher */}
      <Tabs value={activeTab} onValueChange={setActiveTab} className="w-full">
        <TabsList className="grid w-full grid-cols-2 sm:grid-cols-3 lg:grid-cols-4 max-w-[800px] mb-2">
          <TabsTrigger value="table" className="text-[10px] sm:text-xs flex items-center gap-2 font-semibold tracking-wide py-1.5">
            <TableIcon className="w-3 h-3" />
            TABLE VIEW
          </TabsTrigger>
          <TabsTrigger value="json" className="text-[10px] sm:text-xs flex items-center gap-2 font-semibold tracking-wide py-1.5">
            <FileJson className="w-3 h-3" />
            JSON VIEW
          </TabsTrigger>
          {showSummary && (
            <TabsTrigger value="summary" className="text-[10px] sm:text-xs flex items-center gap-2 font-semibold tracking-wide py-1.5">
              <Brain className="w-3 h-3" />
              AI SUMMARY
            </TabsTrigger>
          )}
          {showMerge && (
            <TabsTrigger value="merge" className="text-[10px] sm:text-xs flex items-center gap-2 font-semibold tracking-wide py-1.5">
              <Merge className="w-3 h-3" />
              MERGE SUMMARY
            </TabsTrigger>
          )}
        </TabsList>

        <TabsContent value="table" className="mt-0">
          <TableView data={tableData} title="Extracted Data Grid" maxHeight="450px" />
        </TabsContent>

        <TabsContent value="json" className="mt-0">
          <JsonViewer data={result} title="Raw Extraction Data" maxHeight="450px" />
        </TabsContent>

        {isLossRun && (
          <TabsContent value="summary" className="mt-0">
            <div className="rounded-lg border border-border bg-muted/20 min-h-[200px] p-5">
              {!summaryText && !isSummarizing && (
                <div className="flex flex-col items-center justify-center gap-4 py-10">
                  <div className="p-3 rounded-full bg-primary/10">
                    <Brain className="w-8 h-8 text-primary" />
                  </div>
                  <div className="text-center">
                    <p className="text-sm font-semibold text-foreground mb-1">AI Claims Analysis</p>
                    <p className="text-xs text-muted-foreground">Click Analyze to generate an AI-powered summary of the claims data</p>
                  </div>
                  <Button
                    onClick={handleAnalyze}
                    className="bg-primary hover:bg-primary/90 text-primary-foreground"
                  >
                    <Brain className="w-4 h-4 mr-2" />
                    Analyze
                  </Button>
                </div>
              )}

              {isSummarizing && (
                <div className="flex flex-col items-center justify-center gap-4 py-10">
                  <Loader2 className="w-8 h-8 text-primary animate-spin" />
                  <p className="text-sm text-muted-foreground">Generating AI summary...</p>
                </div>
              )}

              {summaryText && !isSummarizing && (
                <div className="space-y-4">
                  <div className="flex items-center justify-between">
                    <p className="text-xs font-semibold text-primary uppercase tracking-wide flex items-center gap-1.5">
                      <Brain className="w-3.5 h-3.5" />
                      AI Claims Analysis Summary
                    </p>
                    <div className="flex gap-2">
                      <Button variant="outline" size="sm" className="text-xs h-7" onClick={handleCopySummary}>
                        <Copy className="w-3 h-3 mr-1.5" />
                        Copy
                      </Button>
                      <Button variant="ghost" size="sm" className="text-xs h-7 text-muted-foreground" onClick={handleAnalyze}>
                        <RotateCcw className="w-3 h-3 mr-1.5" />
                        Re-analyze
                      </Button>
                    </div>
                  </div>
                  <div className="font-sans text-sm text-foreground/90 leading-relaxed bg-card p-4 rounded-lg border border-border/50 max-h-[400px] overflow-y-auto">
                    {renderMarkdown(summaryText)}
                  </div>
                  <p className="text-[10px] text-muted-foreground italic">
                    * Generated by AI based on extracted claims data. Please verify financial totals.
                  </p>
                </div>
              )}
            </div>
          </TabsContent>
        )}

        {hasMultipleDocs && (
          <TabsContent value="merge" className="mt-0">
            <div className="rounded-lg border border-border bg-muted/20 min-h-[200px] p-5">
              <div className="flex flex-col gap-6">
                {/* Global Merge Actions */}
                <div className="flex flex-col gap-3">
                  <p className="text-[11px] font-bold text-muted-foreground uppercase tracking-wider">Multi-Document Actions</p>
                  <div className="flex gap-2">
                    <Button
                      size="sm"
                      onClick={onDownloadMergedJson}
                      className="bg-stage-done hover:bg-stage-done/90 text-primary-foreground flex-1"
                    >
                      <Merge className="w-4 h-4 mr-2" />
                      Merge JSON
                      <Download className="w-4 h-4 ml-2" />
                    </Button>
                    <Button
                      size="sm"
                      variant="outline"
                      onClick={onDownloadMergedCsv}
                      className="border-stage-done text-stage-done hover:bg-stage-done/5 flex-1"
                    >
                      <Download className="w-4 h-4 mr-2" />
                      Merge CSV
                    </Button>
                  </div>
                </div>

                <div className="h-px bg-border" />

                {/* Merged Analysis Section */}
                {!mergedSummary && !isMerging && (
                  <div className="flex flex-col items-center justify-center gap-4 py-6">
                    <div className="p-3 rounded-full bg-primary/10">
                      <Brain className="w-8 h-8 text-primary" />
                    </div>
                    <div className="text-center">
                      <p className="text-sm font-semibold text-foreground mb-1">Merged AI Analysis</p>
                      <p className="text-xs text-muted-foreground">Analyze all successfully processed documents together</p>
                    </div>
                    <Button
                      onClick={onTriggerMergeAnalysis}
                      className="bg-primary hover:bg-primary/90 text-primary-foreground"
                    >
                      <Brain className="w-4 h-4 mr-2" />
                      Run Merged Analysis
                    </Button>
                  </div>
                )}

                {isMerging && (
                  <div className="flex flex-col items-center justify-center gap-4 py-10">
                    <Loader2 className="w-8 h-8 text-primary animate-spin" />
                    <p className="text-sm text-muted-foreground">Generating merged AI summary...</p>
                  </div>
                )}

                {mergedSummary && !isMerging && (
                  <div className="space-y-4">
                    <div className="flex items-center justify-between">
                      <p className="text-xs font-semibold text-primary uppercase tracking-wide flex items-center gap-1.5">
                        <Brain className="w-3.5 h-3.5" />
                        Comprehensive Merged Analysis
                      </p>
                      <div className="flex gap-2">
                        <Button variant="ghost" size="sm" className="text-xs h-7 text-muted-foreground" onClick={onTriggerMergeAnalysis}>
                          <RotateCcw className="w-3 h-3 mr-1.5" />
                          Re-analyze
                        </Button>
                      </div>
                    </div>
                    <div className="font-sans text-sm text-foreground/90 leading-relaxed bg-card p-4 rounded-lg border border-border/50 max-h-[400px] overflow-y-auto">
                      {renderMarkdown(mergedSummary)}
                    </div>
                    <p className="text-[10px] text-muted-foreground italic">
                      * Consolidated analysis of all uploaded documents.
                    </p>
                  </div>
                )}
              </div>
            </div>
          </TabsContent>
        )}
      </Tabs>
    </div>
  );
};

export default ResultsPanel;
