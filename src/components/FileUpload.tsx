'use client';

import { useState } from 'react';
import { Upload, AlertCircle, CheckCircle2, FileText } from 'lucide-react';

interface FileUploadProps {
  onFilesUploaded: (files: {
    staffTulsi?: File;
    workerTulsi?: File;
    dueVoucher?: File;
    bonusSummary?: File;
    actualPercentage?: File;
    monthWise?: File;
    loanDeduction?: File;
    hrComparison?: File;
  }) => void;
}

export default function FileUpload({ onFilesUploaded }: FileUploadProps) {
  const [files, setFiles] = useState<{
    staffTulsi?: File;
    workerTulsi?: File;
    dueVoucher?: File;
    bonusSummary?: File;
    actualPercentage?: File;
    monthWise?: File;
    loanDeduction?: File;
    hrComparison?: File;
  }>({});

  const matchFile = (filename: string): keyof typeof files | null => {
    const name = filename.toLowerCase();
    
    // Check for each file type with flexible matching
    if (name.includes('staff') && name.includes('tulsi')) return 'staffTulsi';
    if (name.includes('worker') && name.includes('tulsi')) return 'workerTulsi';
    if (name.includes('due') && name.includes('voucher')) return 'dueVoucher';
    if ((name.includes('bonus') && (name.includes('summary') || name.includes('summery'))) 
        && !name.includes('hr') && !name.includes('calculation')) return 'bonusSummary';
    if (name.includes('actual') && name.includes('percentage')) return 'actualPercentage';
    if (name.includes('month') && name.includes('wise')) return 'monthWise';
    if (name.includes('loan') && name.includes('deduction')) return 'loanDeduction';
    if (name.includes('hr') || (name.includes('bonus') && name.includes('calculation'))) return 'hrComparison';
    
    return null;
  };

  const handleFilesChange = (selectedFiles: FileList | null) => {
    if (!selectedFiles) return;

    const newFiles: typeof files = { ...files };

    Array.from(selectedFiles).forEach((file) => {
      const key = matchFile(file.name);
      if (key) {
        newFiles[key] = file;
      }
    });

    setFiles(newFiles);
  };

  const handleProcess = () => {
    if (files.staffTulsi && files.workerTulsi) {
      onFilesUploaded(files);
    }
  };

  const getDisplayName = (key: string): string => {
    const names: Record<string, string> = {
      staffTulsi: 'Staff Tulsi',
      workerTulsi: 'Worker Tulsi',
      dueVoucher: 'Due Voucher',
      bonusSummary: 'Bonus Summary',
      actualPercentage: 'Actual Percentage',
      monthWise: 'Month Wise',
      loanDeduction: 'Loan Deduction',
      hrComparison: 'HR Comparison'
    };
    return names[key] || key;
  };

  return (
    <div className="bg-white p-6 rounded-lg shadow-md">
      <div className="flex items-center mb-4">
        <FileText className="h-6 w-6 mr-2 text-blue-600" />
        <h3 className="text-lg font-semibold">Upload All Files</h3>
      </div>

      <div className="mb-4">
        <label className="block text-sm font-medium mb-2 text-gray-700">
          Select all 8 files at once
        </label>
        <input
          type="file"
          accept=".xlsx,.xls"
          multiple
          onChange={(e) => handleFilesChange(e.target.files)}
          className="w-full px-2 py-2 border rounded text-sm"
        />
      </div>

      <div className="mb-4">
        <h4 className="text-sm font-medium text-gray-700 mb-2">Uploaded Files</h4>
        <div className="space-y-2">
          {Object.entries(files).map(([key, file]) => (
            <div
              key={key}
              className="flex items-center text-green-600 text-sm bg-green-50 rounded px-2 py-1"
            >
              <CheckCircle2 className="h-4 w-4 mr-2 flex-shrink-0" />
              <span className="font-medium">{getDisplayName(key)}:</span>&nbsp;
              <span className="truncate">{file?.name}</span>
            </div>
          ))}
          {Object.keys(files).length === 0 && (
            <div className="text-gray-500 text-sm">No files uploaded yet.</div>
          )}
        </div>
      </div>

      {!files.staffTulsi || !files.workerTulsi ? (
        <div className="flex items-center text-amber-600 text-sm mb-4 p-3 bg-amber-50 rounded-lg">
          <AlertCircle className="h-4 w-4 mr-2 flex-shrink-0" />
          <span>Both Staff Tulsi and Worker Tulsi files are required</span>
        </div>
      ) : (
        <div className="flex items-center text-green-600 text-sm mb-4 p-3 bg-green-50 rounded-lg">
          <CheckCircle2 className="h-4 w-4 mr-2 flex-shrink-0" />
          <span>Required files uploaded. Ready to process!</span>
        </div>
      )}

      <div className="text-xs text-gray-500 mb-4">
        Uploaded: {Object.keys(files).length} of 8 files
      </div>

      <button
        onClick={handleProcess}
        disabled={!files.staffTulsi || !files.workerTulsi}
        className="w-full bg-blue-600 hover:bg-blue-700 disabled:bg-gray-400 disabled:cursor-not-allowed text-white py-2.5 px-4 rounded-lg flex items-center justify-center text-sm font-medium transition-colors shadow-sm"
      >
        <Upload className="h-4 w-4 mr-2" />
        Process Files and Generate Bonus Report
      </button>
    </div>
  );
}