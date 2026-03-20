import { Copy, Check } from "lucide-react";
import { useState } from "react";
import { Button } from "@/components/ui/button";

interface JsonViewerProps {
  data: unknown;
  title?: string;
  maxHeight?: string;
}

const JsonViewer = ({ data, title, maxHeight = "400px" }: JsonViewerProps) => {
  const [copied, setCopied] = useState(false);
  const jsonStr = JSON.stringify(data, null, 2);

  const handleCopy = () => {
    navigator.clipboard.writeText(jsonStr);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  return (
    <div className="rounded-lg overflow-hidden border border-border">
      {title && (
        <div className="flex items-center justify-between px-4 py-2 bg-muted">
          <span className="text-xs font-medium text-muted-foreground">{title}</span>
          <Button variant="ghost" size="sm" className="h-7 text-xs" onClick={handleCopy}>
            {copied ? (
              <><Check className="w-3 h-3 mr-1" /> Copied</>
            ) : (
              <><Copy className="w-3 h-3 mr-1" /> Copy</>
            )}
          </Button>
        </div>
      )}
      <pre
        className="json-viewer p-4 text-xs font-mono overflow-auto whitespace-pre-wrap"
        style={{ maxHeight }}
      >
        {jsonStr}
      </pre>
    </div>
  );
};

export default JsonViewer;
