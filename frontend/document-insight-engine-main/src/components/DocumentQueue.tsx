import { FileText, CheckCircle2, Loader2, Clock, AlertCircle, ChevronDown, ChevronUp } from "lucide-react";
import { useState } from "react";
import type { DocumentFile } from "@/types/extractor";
import ProcessingStages from "./ProcessingStages";
import { STAGE_LABELS } from "@/types/extractor";

interface DocumentQueueProps {
  documents: DocumentFile[];
  activeDocId: string | null;
  onSelectDoc: (id: string) => void;
}

const statusIcon = (doc: DocumentFile) => {
  if (doc.stage === "complete") return <CheckCircle2 className="w-4 h-4 text-stage-done" />;
  if (doc.stage === "error") return <AlertCircle className="w-4 h-4 text-stage-error" />;
  if (doc.stage === "queued") return <Clock className="w-4 h-4 text-badge-queued" />;
  return <Loader2 className="w-4 h-4 text-stage-active animate-spin" />;
};

const statusBadge = (doc: DocumentFile) => {
  if (doc.stage === "complete")
    return <span className="text-[10px] px-2 py-0.5 rounded-full bg-stage-done/15 text-stage-done font-medium">Complete</span>;
  if (doc.stage === "error")
    return <span className="text-[10px] px-2 py-0.5 rounded-full bg-stage-error/15 text-stage-error font-medium">Error</span>;
  if (doc.stage === "queued")
    return <span className="text-[10px] px-2 py-0.5 rounded-full bg-badge-queued/15 text-badge-queued font-medium">Queued</span>;
  return <span className="text-[10px] px-2 py-0.5 rounded-full bg-stage-active/15 text-stage-active font-medium">Processing</span>;
};

const DocumentQueue = ({ documents, activeDocId, onSelectDoc }: DocumentQueueProps) => {
  const [expandedId, setExpandedId] = useState<string | null>(null);

  if (!documents.length) return null;

  return (
    <div className="space-y-2">
      <h3 className="text-sm font-semibold text-foreground flex items-center gap-2">
        <FileText className="w-4 h-4" />
        Document Queue ({documents.length})
      </h3>
      <div className="space-y-1.5 max-h-[500px] overflow-y-auto pr-1">
        {documents.map((doc) => {
          const isExpanded = expandedId === doc.id;
          const isActive = activeDocId === doc.id;
          return (
            <div
              key={doc.id}
              className={`rounded-lg border transition-all duration-200 animate-slide-up ${
                isActive ? "border-primary/30 bg-primary/3" : "border-border bg-card"
              }`}
            >
              <div
                className="flex items-center gap-3 px-3 py-2.5 cursor-pointer hover:bg-muted/50 rounded-lg"
                onClick={() => onSelectDoc(doc.id)}
              >
                {statusIcon(doc)}
                <div className="flex-1 min-w-0">
                  <p className="text-sm font-medium truncate text-foreground">{doc.name}</p>
                  <p className="text-[11px] text-muted-foreground">
                    {(doc.size / 1024 / 1024).toFixed(2)} MB
                    {doc.stage !== "queued" && doc.stage !== "complete" && doc.stage !== "error" && (
                      <> • {STAGE_LABELS[doc.stage]}</>
                    )}
                    {doc.stage === "complete" && doc.result && (
                      <> • {doc.result.claims_count} claims found</>
                    )}
                  </p>
                </div>
                {statusBadge(doc)}
                <button
                  onClick={(e) => {
                    e.stopPropagation();
                    setExpandedId(isExpanded ? null : doc.id);
                  }}
                  className="text-muted-foreground hover:text-foreground p-0.5"
                >
                  {isExpanded ? <ChevronUp className="w-3.5 h-3.5" /> : <ChevronDown className="w-3.5 h-3.5" />}
                </button>
              </div>
              {isExpanded && doc.stage !== "queued" && (
                <div className="px-3 pb-3 border-t border-border/50 pt-2">
                  <ProcessingStages currentStage={doc.stage} stageMessage={doc.stageMessage} />
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
};

export default DocumentQueue;
