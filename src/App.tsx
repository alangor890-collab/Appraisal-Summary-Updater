/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useCallback } from 'react';
import { useDropzone } from 'react-dropzone';
import * as XLSX from 'xlsx';
import { PDFDocument } from 'pdf-lib';
import { GoogleGenAI, GenerateContentResponse } from "@google/genai";
import { 
  FileUp, 
  FileText, 
  Table, 
  CheckCircle2, 
  AlertCircle, 
  Loader2, 
  Download,
  ArrowRight,
  RefreshCcw,
  Clock
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';

// Error Boundary Component
class ErrorBoundary extends React.Component<{ children: React.ReactNode }, { hasError: boolean, error: Error | null }> {
  constructor(props: { children: React.ReactNode }) {
    super(props);
    this.state = { hasError: false, error: null };
  }

  static getDerivedStateFromError(error: Error) {
    return { hasError: true, error };
  }

  render() {
    if (this.state.hasError) {
      return (
        <div className="min-h-screen bg-red-50 flex items-center justify-center p-6">
          <div className="bg-white p-8 rounded-3xl shadow-xl max-w-md w-full border border-red-100">
            <AlertCircle className="w-12 h-12 text-red-500 mb-4" />
            <h2 className="text-2xl font-bold text-red-900 mb-2">Something went wrong</h2>
            <p className="text-red-700 mb-6">{this.state.error?.message || "An unexpected error occurred."}</p>
            <button 
              onClick={() => window.location.reload()}
              className="w-full py-3 bg-red-600 text-white rounded-xl font-medium hover:bg-red-700 transition-colors"
            >
              Reload Application
            </button>
          </div>
        </div>
      );
    }
    return this.props.children;
  }
}

// Initialize Gemini
const getApiKey = () => {
  // In AI Studio, it's injected as process.env.GEMINI_API_KEY
  // In Vercel (Vite), it's usually VITE_GEMINI_API_KEY via import.meta.env
  return import.meta.env.VITE_GEMINI_API_KEY || process.env.GEMINI_API_KEY || '';
};

const ai = new GoogleGenAI({ apiKey: getApiKey() });

interface ExtractionResult {
  employeeNo: string;
  totalScore: number;
}

interface ProcessingStep {
  id: string;
  label: string;
  status: 'pending' | 'loading' | 'success' | 'error';
  message?: string;
}

export default function AppWrapper() {
  return (
    <ErrorBoundary>
      <App />
    </ErrorBoundary>
  );
}

function App() {
  const [pdfFiles, setPdfFiles] = useState<File[]>([]);
  const [excelFile, setExcelFile] = useState<File | null>(null);
  const [excelData, setExcelData] = useState<any[] | null>(null);
  const [assessYear, setAssessYear] = useState(new Date().getFullYear().toString());
  const [isProcessing, setIsProcessing] = useState(false);
  const [steps, setSteps] = useState<ProcessingStep[]>([
    { id: 'ocr', label: 'Extracting data from PDF (Parallel)', status: 'pending' },
    { id: 'match', label: 'Matching Employee No. in Excel', status: 'pending' },
    { id: 'update', label: 'Updating Excel file', status: 'pending' },
  ]);
  const [resultData, setResultData] = useState<ExtractionResult[] | null>(null);
  const [updatedWorkbook, setUpdatedWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [error, setError] = useState<string | null>(null);

  const onDropPdf = useCallback((acceptedFiles: File[]) => {
    setPdfFiles(prev => [...prev, ...acceptedFiles]);
    setError(null);
  }, []);

  const removePdf = (index: number) => {
    setPdfFiles(prev => prev.filter((_, i) => i !== index));
  };

  const onDropExcel = useCallback(async (acceptedFiles: File[]) => {
    const file = acceptedFiles[0];
    setExcelFile(file);
    setError(null);
    
    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer);
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const data = XLSX.utils.sheet_to_json<any>(worksheet);
      setExcelData(data);
    } catch (err) {
      setError("Failed to read Excel file. Please ensure it's a valid .xlsx or .xls file.");
    }
  }, []);

  const { getRootProps: getPdfProps, getInputProps: getPdfInput, isDragActive: isPdfActive } = useDropzone({
    onDrop: onDropPdf,
    accept: { 'application/pdf': ['.pdf'] },
    multiple: true
  });

  const { getRootProps: getExcelProps, getInputProps: getExcelInput, isDragActive: isExcelActive } = useDropzone({
    onDrop: onDropExcel,
    accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'], 'application/vnd.ms-excel': ['.xls'] },
    multiple: false
  });

  const updateStep = (id: string, updates: Partial<ProcessingStep>) => {
    setSteps(prev => prev.map(step => step.id === id ? { ...step, ...updates } : step));
  };

  const [processingTime, setProcessingTime] = useState(0);
  const timerRef = React.useRef<NodeJS.Timeout | null>(null);

  const startTimer = () => {
    setProcessingTime(0);
    timerRef.current = setInterval(() => {
      setProcessingTime(prev => prev + 1);
    }, 1000);
  };

  const stopTimer = () => {
    if (timerRef.current) {
      clearInterval(timerRef.current);
      timerRef.current = null;
    }
  };

  const processFiles = async () => {
    if (pdfFiles.length === 0 || !excelData || !excelFile) return;

    const apiKey = getApiKey();
    if (!apiKey) {
      setError("Gemini API Key is missing. Please set VITE_GEMINI_API_KEY in your Vercel environment variables.");
      return;
    }

    setIsProcessing(true);
    setError(null);
    setSteps(steps.map(s => ({ ...s, status: 'pending', message: undefined })));
    startTimer();

    try {
      // Step 1: OCR and Extraction (Parallelized across all files)
      updateStep('ocr', { status: 'loading' });
      
      const allValidResults: ExtractionResult[] = [];
      let totalPages = 0;

      for (const file of pdfFiles) {
        const pdfArrayBuffer = await file.arrayBuffer();
        const pdfDoc = await PDFDocument.load(pdfArrayBuffer);
        const pageCount = pdfDoc.getPageCount();
        totalPages += pageCount;
        
        updateStep('ocr', { message: `Processing ${file.name} (${pageCount} pages)...` });

        // Process pages in parallel for this file
        const extractionPromises = Array.from({ length: pageCount }).map(async (_, i) => {
          try {
            // Create a new PDF for this single page
            const singlePageDoc = await PDFDocument.create();
            const [copiedPage] = await singlePageDoc.copyPages(pdfDoc, [i]);
            singlePageDoc.addPage(copiedPage);
            const pageBytes = await singlePageDoc.save();
            
            // Safer base64 conversion
            const base64Page = await new Promise<string>((resolve) => {
              const blob = new Blob([pageBytes], { type: 'application/pdf' });
              const reader = new FileReader();
              reader.onloadend = () => {
                const base64 = (reader.result as string).split(',')[1];
                resolve(base64);
              };
              reader.readAsDataURL(blob);
            });

            const response: GenerateContentResponse = await ai.models.generateContent({
              model: "gemini-3-flash-preview",
              contents: [
                {
                  parts: [
                    {
                      inlineData: {
                        mimeType: "application/pdf",
                        data: base64Page
                      }
                    },
                    {
                      text: "This is a 'STAFF APPRAISAL FORM'. Please extract the 'Employee No.' (located at the top right, e.g., 50797) and the 'Total Score' (located at the bottom of the performance rating table, e.g., 82.5). The numbers may be handwritten. Return the result strictly as a JSON object with keys 'employeeNo' (string) and 'totalScore' (number). If this page does not contain an appraisal form or the required fields, return an empty object {}."
                    }
                  ]
                }
              ],
              config: {
                responseMimeType: "application/json"
              }
            });

            let text = response.text || '{}';
            if (text.includes('```')) {
              const match = text.match(/```(?:json)?\s*([\s\S]*?)\s*```/);
              if (match) text = match[1];
            }
            
            const result = JSON.parse(text.trim());
            return result.employeeNo ? (result as ExtractionResult) : null;
          } catch (e) {
            console.error(`Error processing page ${i + 1} of ${file.name}:`, e);
            return null;
          }
        });

        const results = await Promise.all(extractionPromises);
        const validResults = results.filter((r): r is ExtractionResult => r !== null);
        allValidResults.push(...validResults);
      }
      
      if (allValidResults.length === 0) {
        throw new Error("Could not extract any appraisal data from the PDF(s). Please ensure the PDFs contain Employee No. and Total Score fields.");
      }

      setResultData(allValidResults);
      updateStep('ocr', { status: 'success', message: `Extracted ${allValidResults.length} appraisal(s) from ${pdfFiles.length} file(s) (${totalPages} total pages).` });

      // Step 2: Match with current Excel data
      updateStep('match', { status: 'loading' });
      
      // Work on a copy of the data to maintain state correctly
      const data = JSON.parse(JSON.stringify(excelData));

      // Optimization: Create a lookup map for the Excel data
      const idMap = new Map<string, number>();
      const commonHeaders = ['employee no.', 'employee no', 'employee id', 'id', 'employeeno', 'staff no', 'staff id'];
      
      data.forEach((row: any, index: number) => {
        Object.entries(row).forEach(([key, val]) => {
          const normalizedKey = key.toLowerCase().trim();
          const normalizedVal = String(val).trim();
          if (commonHeaders.includes(normalizedKey) || !idMap.has(normalizedVal)) {
            idMap.set(normalizedVal, index);
          }
        });
      });

      let matchedCount = 0;
      const unmatchedIds: string[] = [];

      allValidResults.forEach(result => {
        const targetId = String(result.employeeNo).trim();
        const employeeIndex = idMap.get(targetId);

        if (employeeIndex !== undefined) {
          const scoreKey = Object.keys(data[employeeIndex]).find(k => 
            k.toLowerCase().trim() === 'total score'
          ) || 'Total Score';
          
          data[employeeIndex][scoreKey] = result.totalScore;
          matchedCount++;
        } else {
          unmatchedIds.push(result.employeeNo);
        }
      });

      if (matchedCount === 0) {
        const allKeys = data.length > 0 ? Object.keys(data[0]) : [];
        const detectedColumns = allKeys.join(', ');
        throw new Error(`None of the extracted Employee IDs were found in the Excel file. \n\nDetected columns: [${detectedColumns}].`);
      }

      const matchMessage = unmatchedIds.length > 0 
        ? `Matched ${matchedCount} employees. Unmatched: ${unmatchedIds.join(', ')}`
        : `Successfully matched all ${matchedCount} employees.`;
      
      updateStep('match', { status: 'success', message: matchMessage });

      // Step 3: Update Excel
      updateStep('update', { status: 'loading' });
      
      // Update the session data
      setExcelData(data);

      const workbook = XLSX.utils.book_new();
      
      // Create a worksheet with title and date
      const title = `Staff Appraisal ${assessYear}`;
      const updateDate = `Update Date: ${new Date().toLocaleDateString()}`;
      
      // Prepare headers to find the rightmost column for the date
      const headers = Object.keys(data[0] || {});
      const colCount = Math.max(headers.length, 2);
      
      // Create AOA for the header section
      const headerAOA = [
        [title], // Row 1: Title at A1
        []       // Row 2: Empty spacer
      ];
      
      const worksheet = XLSX.utils.aoa_to_sheet(headerAOA);
      
      // Add the date at the top right corner
      // We use the column count to place it in the last column of the first row
      const dateCellRef = XLSX.utils.encode_cell({ r: 0, c: colCount - 1 });
      XLSX.utils.sheet_add_aoa(worksheet, [[updateDate]], { origin: dateCellRef });

      // Add the actual data starting from row 3 (index 2)
      XLSX.utils.sheet_add_json(worksheet, data, { origin: "A3" });

      XLSX.utils.book_append_sheet(workbook, worksheet, "Staff Appraisal");
      
      setUpdatedWorkbook(workbook);
      updateStep('update', { status: 'success', message: `Updated ${matchedCount} rows in the Excel file.` });

      // Clear processed PDFs to allow next batch
      setPdfFiles([]);

    } catch (err: any) {
      setError(err.message || "An unexpected error occurred during processing.");
      setSteps(prev => prev.map(s => s.status === 'loading' ? { ...s, status: 'error' } : s));
    } finally {
      setIsProcessing(false);
      stopTimer();
    }
  };

  const downloadUpdatedFile = () => {
    if (!updatedWorkbook) return;
    XLSX.writeFile(updatedWorkbook, `Staff_Appraisal_${assessYear}.xlsx`);
  };

  const reset = () => {
    setPdfFiles([]);
    setExcelFile(null);
    setExcelData(null);
    setResultData(null);
    setUpdatedWorkbook(null);
    setError(null);
    setSteps(steps.map(s => ({ ...s, status: 'pending', message: undefined })));
  };

  return (
    <div className="min-h-screen bg-[#f5f5f5] text-[#1a1a1a] font-sans p-6 md:p-12">
      <div className="max-w-4xl mx-auto">
        {/* Header */}
        <header className="mb-12">
          <h1 className="text-4xl font-light tracking-tight mb-2">Appraisal Automator</h1>
          <p className="text-[#9e9e9e] text-lg">AI-powered PDF extraction and Excel synchronization.</p>
        </header>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
          {/* Upload Section */}
          <section className="space-y-6">
            <div className="bg-white p-8 rounded-[24px] shadow-sm border border-black/5">
              <h2 className="text-xs uppercase tracking-widest font-semibold text-[#9e9e9e] mb-6">1. Configure & Upload</h2>
              
              <div className="space-y-4">
                {/* Assess Year Input */}
                <div className="mb-6">
                  <label className="text-xs font-semibold text-[#1a1a1a] mb-2 block">Assessment Year</label>
                  <input 
                    type="text" 
                    value={assessYear}
                    onChange={(e) => setAssessYear(e.target.value)}
                    placeholder="e.g. 2025"
                    className="w-full px-4 py-3 rounded-xl border border-[#e5e5e5] focus:border-[#1a1a1a] focus:ring-1 focus:ring-[#1a1a1a] outline-none transition-all text-sm"
                  />
                </div>

                {/* PDF Dropzone */}
                <div 
                  {...getPdfProps()} 
                  className={`border-2 border-dashed rounded-2xl p-6 transition-all cursor-pointer flex flex-col items-center justify-center text-center
                    ${pdfFiles.length > 0 ? 'border-green-200 bg-green-50/30' : 'border-[#e5e5e5] hover:border-[#9e9e9e]'}
                    ${isPdfActive ? 'border-blue-400 bg-blue-50/30' : ''}`}
                >
                  <input {...getPdfInput()} />
                  <FileText className={`w-8 h-8 mb-3 ${pdfFiles.length > 0 ? 'text-green-500' : 'text-[#9e9e9e]'}`} />
                  {pdfFiles.length > 0 ? (
                    <div className="w-full space-y-2">
                      <span className="text-sm font-medium block">{pdfFiles.length} PDF(s) selected</span>
                      <div className="max-h-32 overflow-y-auto space-y-1 px-2">
                        {pdfFiles.map((file, idx) => (
                          <div key={idx} className="flex items-center justify-between text-xs bg-white/50 p-1 rounded border border-black/5">
                            <span className="truncate flex-grow text-left">{file.name}</span>
                            <button 
                              onClick={(e) => { e.stopPropagation(); removePdf(idx); }}
                              className="ml-2 text-red-500 hover:text-red-700"
                            >
                              ✕
                            </button>
                          </div>
                        ))}
                      </div>
                    </div>
                  ) : (
                    <span className="text-sm text-[#9e9e9e]">Drop PDF Appraisal Form(s)</span>
                  )}
                </div>

                {/* Excel Dropzone */}
                <div 
                  {...getExcelProps()} 
                  className={`border-2 border-dashed rounded-2xl p-6 transition-all cursor-pointer flex flex-col items-center justify-center text-center
                    ${excelFile ? 'border-green-200 bg-green-50/30' : 'border-[#e5e5e5] hover:border-[#9e9e9e]'}
                    ${isExcelActive ? 'border-blue-400 bg-blue-50/30' : ''}`}
                >
                  <input {...getExcelInput()} />
                  <Table className={`w-8 h-8 mb-3 ${excelFile ? 'text-green-500' : 'text-[#9e9e9e]'}`} />
                  {excelFile ? (
                    <span className="text-sm font-medium truncate max-w-full">{excelFile.name}</span>
                  ) : (
                    <span className="text-sm text-[#9e9e9e]">Drop Staff List Excel</span>
                  )}
                </div>
              </div>

              <button
                onClick={processFiles}
                disabled={pdfFiles.length === 0 || !excelData || isProcessing}
                className={`w-full mt-8 py-4 rounded-2xl font-medium transition-all flex items-center justify-center gap-2
                  ${pdfFiles.length === 0 || !excelData || isProcessing 
                    ? 'bg-[#e5e5e5] text-[#9e9e9e] cursor-not-allowed' 
                    : 'bg-[#1a1a1a] text-white hover:bg-black shadow-lg hover:shadow-xl active:scale-[0.98]'}`}
              >
                {isProcessing ? (
                  <>
                    <Loader2 className="w-5 h-5 animate-spin" />
                    Processing...
                  </>
                ) : (
                  <>
                    Start Automation
                    <ArrowRight className="w-5 h-5" />
                  </>
                )}
              </button>
            </div>
          </section>

          {/* Status Section */}
          <section className="space-y-6">
            <div className="bg-white p-8 rounded-[24px] shadow-sm border border-black/5 min-h-[400px] flex flex-col">
              <h2 className="text-xs uppercase tracking-widest font-semibold text-[#9e9e9e] mb-6 flex justify-between items-center">
                2. Automation Status
                {isProcessing && (
                  <span className="font-mono text-blue-500 lowercase tracking-normal">
                    {Math.floor(processingTime / 60)}:{(processingTime % 60).toString().padStart(2, '0')}
                  </span>
                )}
              </h2>
              
              <div className="flex-grow space-y-6">
                {steps.map((step, index) => (
                  <div key={step.id} className="flex gap-4">
                    <div className="flex flex-col items-center">
                      <div className={`w-8 h-8 rounded-full flex items-center justify-center border-2 transition-all
                        ${step.status === 'success' ? 'bg-green-500 border-green-500 text-white' : 
                          step.status === 'loading' ? 'border-blue-500 text-blue-500' : 
                          step.status === 'error' ? 'bg-red-500 border-red-500 text-white' : 
                          'border-[#e5e5e5] text-[#e5e5e5]'}`}
                      >
                        {step.status === 'success' ? <CheckCircle2 className="w-5 h-5" /> : 
                         step.status === 'loading' ? <Loader2 className="w-5 h-5 animate-spin" /> : 
                         step.status === 'error' ? <AlertCircle className="w-5 h-5" /> : 
                         <span className="text-xs font-bold">{index + 1}</span>}
                      </div>
                      {index < steps.length - 1 && (
                        <div className={`w-0.5 h-full my-1 rounded-full ${steps[index + 1].status !== 'pending' ? 'bg-green-200' : 'bg-[#f5f5f5]'}`} />
                      )}
                    </div>
                    <div className="pb-6">
                      <p className={`font-medium ${step.status === 'pending' ? 'text-[#9e9e9e]' : 'text-[#1a1a1a]'}`}>
                        {step.label}
                      </p>
                      {step.message && (
                        <p className="text-sm text-[#9e9e9e] mt-1">{step.message}</p>
                      )}
                    </div>
                  </div>
                ))}
              </div>

              <AnimatePresence>
                {error && (
                  <motion.div 
                    initial={{ opacity: 0, y: 10 }}
                    animate={{ opacity: 1, y: 0 }}
                    className="mt-4 p-4 bg-red-50 border border-red-100 rounded-xl flex gap-3 items-start"
                  >
                    <AlertCircle className="w-5 h-5 text-red-500 shrink-0 mt-0.5" />
                    <p className="text-sm text-red-700">{error}</p>
                  </motion.div>
                )}

                {updatedWorkbook && (
                  <motion.div 
                    initial={{ opacity: 0, scale: 0.95 }}
                    animate={{ opacity: 1, scale: 1 }}
                    className="mt-6 space-y-3"
                  >
                    <button
                      onClick={downloadUpdatedFile}
                      className="w-full py-4 bg-green-500 text-white rounded-2xl font-medium hover:bg-green-600 transition-all flex items-center justify-center gap-2 shadow-lg shadow-green-200"
                    >
                      <Download className="w-5 h-5" />
                      Download Updated Excel
                    </button>
                    <button
                      onClick={reset}
                      className="w-full py-3 text-[#9e9e9e] hover:text-[#1a1a1a] transition-all flex items-center justify-center gap-2 text-sm"
                    >
                      <RefreshCcw className="w-4 h-4" />
                      Start New Process
                    </button>
                  </motion.div>
                )}
              </AnimatePresence>
            </div>
          </section>
        </div>

        {/* Footer info */}
        <footer className="mt-12 text-center text-[#9e9e9e] text-sm">
          <p>Powered by Gemini 3 Flash for high-precision OCR and data extraction.</p>
        </footer>
      </div>
    </div>
  );
}
