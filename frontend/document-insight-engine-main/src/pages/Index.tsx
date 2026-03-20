import AppHeader from "@/components/AppHeader";
import UploadArea from "@/components/UploadArea";
import DocumentQueue from "@/components/DocumentQueue";
import ResultsPanel from "@/components/ResultsPanel";
import MergeJsonButton from "@/components/MergeJsonButton";
import ProcessingStages from "@/components/ProcessingStages";
import { useDocumentProcessor } from "@/hooks/useDocumentProcessor";
import { FileText } from "lucide-react";

const Index = () => {
  const { documents, activeDocId, selectedDoc, addFiles, setActiveDocId } =
    useDocumentProcessor();

  const hasDocuments = documents.length > 0;
  const isAnyProcessing = documents.some(
    (d) => d.stage !== "queued" && d.stage !== "complete" && d.stage !== "error"
  );

  return (
    <div className="min-h-screen bg-background">
      <AppHeader />

      <main className="max-w-7xl mx-auto px-4 sm:px-8 py-8">
        {/* Upload Section */}
        <section className="mb-8">
          <UploadArea onFilesSelected={addFiles} />
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
                    <h3 className="text-sm font-semibold text-foreground mb-3">
                      Processing: {selectedDoc.name}
                    </h3>
                    <ProcessingStages
                      currentStage={selectedDoc.stage}
                      stageMessage={selectedDoc.stageMessage}
                    />
                  </div>
                )}

              {/* Merge button */}
              <MergeJsonButton documents={documents} />
            </aside>

            {/* Right panel: Results */}
            <section className="lg:col-span-8">
              <div className="p-6 rounded-xl border border-border bg-card min-h-[400px]">
                <h2 className="text-base font-semibold text-foreground mb-4 flex items-center gap-2">
                  <FileText className="w-4 h-4 text-primary" />
                  {selectedDoc
                    ? `Results — ${selectedDoc.name}`
                    : "Extraction Results"}
                </h2>
                <ResultsPanel document={selectedDoc} />
              </div>
            </section>
          </div>
        )}
      </main>

      {/* Footer */}
      <footer className="text-center py-6 text-xs text-muted-foreground">
        Insurance Form Extractor • AI-Powered PDF Processing
      </footer>
    </div>
  );
};

export default Index;
