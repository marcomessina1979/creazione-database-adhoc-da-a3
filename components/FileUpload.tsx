
import React, { useState, useCallback } from 'react';
import { Spinner } from './Spinner';

interface FileUploadProps {
  id: string;
  onFileSelect: (file: File | null) => void;
  acceptedFileType: string;
}

export const FileUpload: React.FC<FileUploadProps> = ({ id, onFileSelect, acceptedFileType }) => {
  const [fileName, setFileName] = useState<string | null>(null);
  const [isDragging, setIsDragging] = useState<boolean>(false);
  const [isUploading, setIsUploading] = useState<boolean>(false);

  const handleFileChange = (files: FileList | null) => {
    if (files && files.length > 0) {
      const file = files[0];
      setFileName(file.name);
      setIsUploading(true);
      // Simulate processing for better UX, prevents UI lag perception on large files
      setTimeout(() => {
        onFileSelect(file);
        setIsUploading(false);
      }, 750);
    } else {
      setFileName(null);
      onFileSelect(null);
    }
  };

  const onDragEnter = useCallback((e: React.DragEvent<HTMLLabelElement>) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(true);
  }, []);

  const onDragLeave = useCallback((e: React.DragEvent<HTMLLabelElement>) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
  }, []);

  const onDragOver = useCallback((e: React.DragEvent<HTMLLabelElement>) => {
    e.preventDefault();
    e.stopPropagation();
  }, []);

  const onDrop = useCallback((e: React.DragEvent<HTMLLabelElement>) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
    if(isUploading) return;
    handleFileChange(e.dataTransfer.files);
  }, [isUploading]);

  return (
      <label
        htmlFor={id}
        className={`relative flex flex-col items-center justify-center w-full h-32 px-4 transition-all duration-300 bg-white border-2 border-dashed rounded-lg appearance-none group hover:border-blue-500 focus:outline-none ${isDragging ? 'border-blue-600 bg-blue-50' : 'border-gray-300'} ${isUploading ? 'cursor-wait bg-gray-50' : 'cursor-pointer hover:bg-gray-50'}`}
        onDragEnter={onDragEnter}
        onDragLeave={onDragLeave}
        onDragOver={onDragOver}
        onDrop={onDrop}
      >
        {isUploading ? (
            <div className="flex flex-col items-center text-blue-600">
              <Spinner />
              <span className="font-medium mt-2">Caricamento...</span>
            </div>
        ) : (
            <div className="text-center">
              <div className="flex flex-col items-center justify-center space-y-2">
                  <svg xmlns="http://www.w3.org/2000/svg" className="w-8 h-8 text-gray-400 group-hover:text-blue-600 transition-colors" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth="2">
                      <path strokeLinecap="round" strokeLinejoin="round" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
                  </svg>
                  <span className="font-medium text-gray-600 text-sm">
                      {fileName ? <span className="break-all font-semibold text-blue-800">{fileName}</span> : <>Trascina o <span className="text-blue-600 underline">cerca</span> un file</>}
                  </span>
              </div>
            </div>
        )}
        <input 
            type="file" 
            id={id} 
            name={id}
            className="absolute inset-0 w-full h-full opacity-0 cursor-pointer disabled:cursor-wait" 
            accept={acceptedFileType}
            onChange={(e) => handleFileChange(e.target.files)}
            disabled={isUploading}
        />
      </label>
  );
};
