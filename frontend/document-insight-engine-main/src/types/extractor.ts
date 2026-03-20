export type ProcessingStage =
  | "queued"
  | "rotation_check"
  | "text_extraction"
  | "schema_extraction"
  | "policy_detection"
  | "claim_extraction"
  | "validation"
  | "complete"
  | "error";

export const STAGE_LABELS: Record<ProcessingStage, string> = {
  queued: "Queued",
  rotation_check: "Checking Rotation",
  text_extraction: "Extracting Text",
  schema_extraction: "Schema Extraction",
  policy_detection: "Policy Detection & Chunking",
  claim_extraction: "Extracting Claims",
  validation: "Validating",
  complete: "Complete",
  error: "Error",
};

export const STAGE_ORDER: ProcessingStage[] = [
  "rotation_check",
  "text_extraction",
  "schema_extraction",
  "policy_detection",
  "claim_extraction",
  "validation",
  "complete",
];

export interface ClaimData {
  claim_number: string;
  claimant_name?: string;
  date_of_loss?: string;
  status?: string;
  [key: string]: unknown;
}

export interface ExtractionResult {
  claims: any[];
  [key: string]: any;
}

export interface DocumentMetadata {
  insurer: string;
  format: string;
  confidence: number;
  claims_count: number;
}

export interface DocumentFile {
  id: string;
  file: File;
  name: string;
  size: number;
  stage: ProcessingStage;
  stageMessage: string;
  progress: number;
  result: ExtractionResult | null;
  metadata?: DocumentMetadata;
  error: string | null;
  startedAt: number | null;
  completedAt: number | null;
}
