import { Download, Merge } from "lucide-react";
import { Button } from "@/components/ui/button";
import type { DocumentFile } from "@/types/extractor";

interface MergeJsonButtonProps {
  documents: DocumentFile[];
}

const MergeJsonButton = ({ documents }: MergeJsonButtonProps) => {
  const completedDocs = documents.filter((d) => d.stage === "complete" && d.result);

  if (completedDocs.length < 2) return null;

  const getAllClaims = () => {
    return completedDocs.flatMap((d) => d.result?.claims || []);
  };

  const handleDownloadJson = () => {
    const claims = getAllClaims();
    const blob = new Blob([JSON.stringify(claims, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = globalThis.document.createElement("a");
    a.href = url;
    a.download = `merged_claims_${Date.now()}.json`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const handleDownloadCsv = () => {
    const claims = getAllClaims();
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
    a.download = `merged_claims_${Date.now()}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  };

  return (
    <div className="flex gap-2">
      <Button
        onClick={handleDownloadJson}
        className="bg-stage-done hover:bg-stage-done/90 text-primary-foreground"
      >
        <Merge className="w-4 h-4 mr-2" />
        Merge to JSON
        <Download className="w-4 h-4 ml-2" />
      </Button>
      <Button
        variant="outline"
        onClick={handleDownloadCsv}
        className="border-stage-done text-stage-done hover:bg-stage-done/5"
      >
        <Download className="w-4 h-4 mr-2" />
        Merge to CSV
      </Button>
    </div>
  );
};

export default MergeJsonButton;
