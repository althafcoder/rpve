import { Download, FileJson, BarChart3, Building2 } from "lucide-react";
import { Button } from "@/components/ui/button";
import type { DocumentFile } from "@/types/extractor";
import JsonViewer from "./JsonViewer";

interface ResultsPanelProps {
  document: DocumentFile | null;
}

const ResultsPanel = ({ document }: ResultsPanelProps) => {
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
      <div className="p-6 bg-destructive/5 rounded-lg border border-destructive/20">
        <p className="text-sm text-destructive font-medium">Processing Error</p>
        <p className="text-xs text-muted-foreground mt-1">{document.error}</p>
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

  const handleDownloadJson = () => {
    const blob = new Blob([JSON.stringify(result, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = globalThis.document.createElement("a");
    a.href = url;
    a.download = `${document.name.replace(".pdf", "")}_extracted.json`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const handleDownloadCsv = () => {
    if (!result?.claims || !Array.isArray(result.claims)) return;

    const claims = result.claims;
    if (claims.length === 0) return;

    // Get all unique keys for headers
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
    const a = globalThis.document.createElement("a");
    a.href = url;
    a.download = `${document.name.replace(".pdf", "")}_extracted.csv`;
    a.click();
    URL.revokeObjectURL(url);
  };

  return (
    <div className="space-y-4 animate-slide-up">
      {/* Summary cards */}
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
            <span className="text-[11px] text-muted-foreground font-medium">Claims Found</span>
          </div>
          <p className="text-sm font-semibold text-foreground">{metadata?.claims_count || result?.claims?.length || 0}</p>
        </div>
        <div className="p-3 rounded-lg bg-muted/50 border border-border">
          <div className="flex items-center gap-2 mb-1">
            <FileJson className="w-3.5 h-3.5 text-primary" />
            <span className="text-[11px] text-muted-foreground font-medium">Confidence</span>
          </div>
          <p className="text-sm font-semibold text-foreground">{metadata?.confidence ? `${metadata.confidence}%` : "N/A"}</p>
        </div>
      </div>

      {/* Actions */}
      <div className="flex gap-2">
        <Button size="sm" onClick={handleDownloadJson}>
          <FileJson className="w-4 h-4 mr-2" />
          Download JSON
        </Button>
        <Button size="sm" variant="outline" onClick={handleDownloadCsv}>
          <Download className="w-4 h-4 mr-2" />
          Download CSV
        </Button>
      </div>

      {/* JSON Viewer */}
      <JsonViewer data={result} title="Raw Extraction Data" maxHeight="350px" />
    </div>
  );
};

export default ResultsPanel;
