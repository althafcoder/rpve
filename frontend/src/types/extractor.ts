export type ProcessingStage =
  | "queued"
  | "classification"
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
  classification: "Intelligent Classification",
  rotation_check: "Checking Rotation",
  text_extraction: "Extracting Text",
  schema_extraction: "Schema Extraction",
  policy_detection: "Policy Detection & Chunking",
  claim_extraction: "Extracting Data",
  validation: "Validating",
  complete: "Complete",
  error: "Error",
};

export const STAGE_ORDER: ProcessingStage[] = [
  "classification",
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
  claims_count?: number;
  total_value?: number;
  documentType?: "INSURANCE" | "INSURANCE_CLAIMS" | "INVOICE" | "VENDOR_INVOICE" | "WORK_COMPENSATION" | "IDENTIFICATION" | "UNKNOWN";
  work_comp_metadata?: {
    form_type: string;
    total_premium: number;
    applicant_name: string;
    wc_states: string[];
  } | null;
}

export interface DocumentFile {
  id: string;
  file: File | File[];
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
  excelPath?: string; // Path to Excel file from backend
  jsonPath?: string;  // Path to JSON file from backend
}
