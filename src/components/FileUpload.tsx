'use client';

import { useState } from 'react';
import { Upload, File, AlertCircle, CheckCircle2 } from 'lucide-react';

interface FileUploadProps {
  onFilesUploaded: (files: { staff?: File; worker?: File; hrComparison?: File }) => void;
}

export default function FileUpload({ onFilesUploaded }: FileUploadProps) {
  const [files, setFiles] = useState<{
    staff?: File;
    worker?: File;
    hrComparison?: File;
  }>({});

  const handleFileChange = (type: 'staff' | 'worker' | 'hrComparison', file: File | null) => {
    const newFiles = { ...files };
    if (file) {
      newFiles[type] = file;
    } else {
      delete newFiles[type];
    }
    setFiles(newFiles);
  };

  const handleProcess = () => {
    if (files.staff && files.worker) {
      onFilesUploaded(files);
    }
  };

  return (
    <div className="bg-white p-4 rounded shadow-sm">
      <h3 className="text-lg font-semibold mb-3">Upload Files</h3>
      
      <div className="grid grid-cols-1 md:grid-cols-2 gap-3 mb-3">
        {/* Staff File */}
        <div>
          <label className="block text-sm font-medium mb-1">Staff File *</label>
          <input
            type="file"
            accept=".xlsx,.xls"
            onChange={(e) => handleFileChange('staff', e.target.files?.[0] || null)}
            className="w-full px-2 py-1.5 border rounded text-sm"
          />
          {files.staff && (
            <div className="flex items-center mt-1 text-green-600 text-xs">
              <CheckCircle2 className="h-3 w-3 mr-1" />
              {files.staff.name}
            </div>
          )}
        </div>

        {/* Worker File */}
        <div>
          <label className="block text-sm font-medium mb-1">Worker File *</label>
          <input
            type="file"
            accept=".xlsx,.xls"
            onChange={(e) => handleFileChange('worker', e.target.files?.[0] || null)}
            className="w-full px-2 py-1.5 border rounded text-sm"
          />
          {files.worker && (
            <div className="flex items-center mt-1 text-green-600 text-xs">
              <CheckCircle2 className="h-3 w-3 mr-1" />
              {files.worker.name}
            </div>
          )}
        </div>
      </div>

      {/* HR Comparison File */}
      <div className="mb-3">
        <label className="block text-sm font-medium mb-1">HR Comparison (Optional)</label>
        <input
          type="file"
          accept=".xlsx,.xls"
          onChange={(e) => handleFileChange('hrComparison', e.target.files?.[0] || null)}
          className="w-full px-2 py-1.5 border rounded text-sm"
        />
        {files.hrComparison && (
          <div className="flex items-center mt-1 text-green-600 text-xs">
            <CheckCircle2 className="h-3 w-3 mr-1" />
            {files.hrComparison.name}
          </div>
        )}
      </div>

      {!files.staff || !files.worker ? (
        <div className="flex items-center text-amber-600 text-sm mb-3">
          <AlertCircle className="h-4 w-4 mr-1" />
          Both Staff and Worker files required
        </div>
      ) : null}

      <button
        onClick={handleProcess}
        disabled={!files.staff || !files.worker}
        className="w-full bg-blue-600 hover:bg-blue-700 disabled:bg-gray-400 text-white py-2 px-3 rounded flex items-center justify-center text-sm"
      >
        <Upload className="h-4 w-4 mr-1" />
        Process Files
      </button>
    </div>
  );
}
