/* eslint-disable @typescript-eslint/explicit-function-return-type */
import * as React from "react";
import { Icon } from "@fluentui/react/lib/Icon";

interface DropZoneUploaderProps {
  onFilesSelected: (files: File[]) => void;
}

const DropZoneUploader: React.FC<DropZoneUploaderProps> = ({ onFilesSelected }) => {
  const fileInputRef = React.useRef<HTMLInputElement>(null);

  const handleDrop = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    if (e.dataTransfer.files.length > 0) {
      onFilesSelected(Array.from(e.dataTransfer.files));
    }
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      onFilesSelected(Array.from(e.target.files));
    }
  };

  return (
    <div
      onDragOver={(e) => e.preventDefault()}
      onDrop={handleDrop}
      onClick={() => fileInputRef.current?.click()}
      style={{
        border: "2px dashed #ccc",
        borderRadius: 6,
        padding: 24,
        textAlign: "center",
        backgroundColor: "#fafafa",
        cursor: "pointer",
        color: "#999",
      }}
    >
      <Icon iconName="Attach" style={{ fontSize: 28 }} />
      <p style={{ margin: 8 }}>Drop your file(s) here or click to upload</p>
      <input
        type="file"
        ref={fileInputRef}
        style={{ display: "none" }}
        onChange={handleFileChange}
        multiple
      />
    </div>
  );
};

export default DropZoneUploader;
