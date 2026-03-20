import { Download, Merge, RotateCcw, Brain, Loader2 } from "lucide-react";
import { Button } from "@/components/ui/button";

interface MergeJsonButtonProps {
  completedDocsCount: number;
  isMergeable: boolean;
  sharedType?: string;
  onDownloadJson: () => void;
  onDownloadCsv: () => void;
  onTriggerMergeAnalysis: () => void;
  isAnalyzing: boolean;
}

const MergeJsonButton = ({
  completedDocsCount,
  isMergeable,
  sharedType,
  onDownloadJson,
  onDownloadCsv,
  onTriggerMergeAnalysis,
  isAnalyzing
}: MergeJsonButtonProps) => {

  if (completedDocsCount < 2 || !isMergeable) return null;

  const typeLabel = sharedType === "WORK_COMPENSATION" ? "Rating Data" : (sharedType === "INSURANCE" ? "Insurance" : "JSON");
  const mergeLabel = `Merge ${completedDocsCount} ${typeLabel}${completedDocsCount > 1 && sharedType !== "WORK_COMPENSATION" ? '' : ''}`;
  // Let's make it cleaner
  const finalMergeLabel = sharedType === "WORK_COMPENSATION"
    ? `Merge ${completedDocsCount} Rating Table(s)`
    : `Merge ${completedDocsCount} ${typeLabel} Docs`;

  return (
    <div className="flex flex-col gap-3 p-4 rounded-xl border border-dashed border-border bg-muted/20">
      <div className="flex flex-col gap-2">
        <div className="flex items-center justify-between px-1">
          <p className="text-[10px] font-bold text-muted-foreground uppercase tracking-widest">Batch Operations</p>
          {!isMergeable && (
            <span className="text-[9px] font-bold text-amber-500 bg-amber-500/10 px-1.5 py-0.5 rounded leading-none">Mixed Types</span>
          )}
        </div>
        <div className="flex gap-2">
          <Button
            onClick={onDownloadJson}
            disabled={!isMergeable}
            className="bg-stage-done hover:bg-stage-done/90 text-primary-foreground flex-1 h-9 disabled:opacity-50 disabled:bg-muted"
          >
            <Merge className="w-4 h-4 mr-2" />
            {finalMergeLabel}
          </Button>
          <Button
            variant="outline"
            onClick={onDownloadCsv}
            disabled={!isMergeable}
            className="border-stage-done text-stage-done hover:bg-stage-done/10 flex-1 h-9 disabled:opacity-30"
          >
            <Download className="w-4 h-4 mr-2" />
            Merge CSV
          </Button>
        </div>
      </div>

      <Button
        onClick={onTriggerMergeAnalysis}
        disabled={isAnalyzing || !isMergeable}
        className="w-full bg-primary hover:bg-primary/90 text-primary-foreground h-10 shadow-sm disabled:opacity-50"
      >
        {isAnalyzing ? (
          <Loader2 className="w-4 h-4 mr-2 animate-spin" />
        ) : (
          <Brain className="w-4 h-4 mr-2" />
        )}
        Merge Analysis (AI Summary)
      </Button>

      <div className="pt-2 mt-1 border-t border-border/50">
        <Button
          variant="ghost"
          onClick={() => window.location.reload()}
          className="w-full h-8 text-xs text-muted-foreground hover:text-destructive hover:bg-destructive/5"
          title="Clear all and start over"
        >
          <RotateCcw className="w-3.5 h-3.5 mr-2" />
          Clear Workspace & Reset
        </Button>
      </div>
    </div>
  );
};

export default MergeJsonButton;