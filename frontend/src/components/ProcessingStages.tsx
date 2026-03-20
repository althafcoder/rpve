import { Check, Loader2, Circle, AlertCircle } from "lucide-react";
import { STAGE_ORDER, STAGE_LABELS, type ProcessingStage } from "@/types/extractor";

interface ProcessingStagesProps {
  currentStage: ProcessingStage;
  stageMessage: string;
}

const ProcessingStages = ({ currentStage, stageMessage }: ProcessingStagesProps) => {
  const currentIdx = STAGE_ORDER.indexOf(currentStage);
  const isError = currentStage === "error";

  return (
    <div className="space-y-1.5">
      {STAGE_ORDER.map((stage, i) => {
        const isDone = currentIdx > i || currentStage === "complete";
        const isActive = currentIdx === i && !isError;
        const isPending = currentIdx < i && !isError;

        return (
          <div
            key={stage}
            className={`flex items-center gap-3 px-3 py-1.5 rounded-md text-sm transition-all duration-300 ${
              isActive ? "bg-primary/8 font-medium" : ""
            }`}
          >
            <div className="flex-shrink-0">
              {isDone ? (
                <Check className="w-4 h-4 text-stage-done" />
              ) : isActive ? (
                <Loader2 className="w-4 h-4 text-stage-active animate-spin" />
              ) : isError && currentIdx === i ? (
                <AlertCircle className="w-4 h-4 text-stage-error" />
              ) : (
                <Circle className="w-3.5 h-3.5 text-stage-pending" />
              )}
            </div>
            <span
              className={`${
                isDone
                  ? "text-stage-done"
                  : isActive
                  ? "text-foreground"
                  : "text-muted-foreground"
              }`}
            >
              {STAGE_LABELS[stage]}
            </span>
            {isActive && stageMessage && (
              <span className="text-xs text-muted-foreground ml-auto truncate max-w-[200px]">
                {stageMessage}
              </span>
            )}
          </div>
        );
      })}
    </div>
  );
};

export default ProcessingStages;
