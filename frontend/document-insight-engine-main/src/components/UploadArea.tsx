import { useCallback, useState, useRef } from "react";
import { Upload, FileUp } from "lucide-react";
import { Button } from "@/components/ui/button";

interface UploadAreaProps {
  onFilesSelected: (files: File[]) => void;
  disabled?: boolean;
}

const UploadArea = ({ onFilesSelected, disabled }: UploadAreaProps) => {
  const [isDragOver, setIsDragOver] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  const handleDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      setIsDragOver(false);
      if (disabled) return;
      const allowedTypes = [
        "application/pdf",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "application/vnd.ms-excel"
      ];
      const files = Array.from(e.dataTransfer.files).filter(
        (f) => allowedTypes.includes(f.type) || f.name.endsWith(".xlsx") || f.name.endsWith(".xls")
      );
      if (files.length) onFilesSelected(files);
    },
    [onFilesSelected, disabled]
  );

  const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = Array.from(e.target.files || []);
    if (files.length) onFilesSelected(files);
    e.target.value = "";
  };

  return (
    <div
      onDragOver={(e) => {
        e.preventDefault();
        if (!disabled) setIsDragOver(true);
      }}
      onDragLeave={() => setIsDragOver(false)}
      onDrop={handleDrop}
      onClick={() => !disabled && inputRef.current?.click()}
      className={`
        relative border-2 border-dashed rounded-xl p-12 text-center cursor-pointer transition-all duration-300
        ${isDragOver ? "gradient-upload-hover border-upload-hover scale-[1.01]" : "border-upload-border bg-upload-bg hover:border-muted-foreground"}
        ${disabled ? "opacity-50 cursor-not-allowed" : ""}
      `}
    >
      <input
        ref={inputRef}
        type="file"
        accept=".pdf,.xlsx,.xls"
        multiple
        className="hidden"
        onChange={handleChange}
      />
      <div className="flex flex-col items-center gap-4">
        <div className={`w-16 h-16 rounded-full flex items-center justify-center transition-colors ${isDragOver ? "bg-primary/10" : "bg-muted"}`}>
          {isDragOver ? (
            <FileUp className="w-8 h-8 text-primary" />
          ) : (
            <Upload className="w-8 h-8 text-muted-foreground" />
          )}
        </div>
        <div>
          <p className="text-lg font-semibold text-foreground">
            {isDragOver ? "Drop files here" : "Drag & Drop Files"}
          </p>
          <p className="text-sm text-muted-foreground mt-1">
            or click to browse • Supports multiple files • PDF & Excel (max 50MB each)
          </p>
        </div>
        <Button
          variant="default"
          size="sm"
          className="mt-2"
          onClick={(e) => {
            e.stopPropagation();
            inputRef.current?.click();
          }}
          disabled={disabled}
        >
          <Upload className="w-4 h-4 mr-2" />
          Select Files
        </Button>
      </div>
    </div>
  );
};

export default UploadArea;
