"use client";
import { useState, useCallback } from "react";
import { motion, AnimatePresence } from "framer-motion";
import { UploadCloud, FileText, Loader2 } from "lucide-react";

export default function DragDropZone({ onUpload }: { onUpload: (file: File) => void }) {
  const [isDragging, setIsDragging] = useState(false);
  const [loading, setLoading] = useState(false);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    if (e.dataTransfer.files?.[0]) {
      setLoading(true);
      onUpload(e.dataTransfer.files[0]);
    }
  }, [onUpload]);

  return (
    <motion.div
      animate={{
        scale: isDragging ? 1.02 : 1,
        borderColor: isDragging ? "#3b82f6" : "#334155",
        backgroundColor: isDragging ? "rgba(59, 130, 246, 0.05)" : "rgba(30, 41, 59, 0.5)",
      }}
      onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
      onDragLeave={() => setIsDragging(false)}
      onDrop={handleDrop}
      className="relative flex flex-col items-center justify-center p-16 border-2 border-dashed rounded-3xl backdrop-blur-md cursor-pointer transition-colors"
    >
      <AnimatePresence mode="wait">
        {loading ? (
          <motion.div key="loading" initial={{ opacity: 0 }} animate={{ opacity: 1 }} className="flex flex-col items-center">
            <Loader2 className="w-12 h-12 text-blue-500 animate-spin mb-4" />
            <p className="text-blue-400 font-medium tracking-wide">Processing Thermodynamics...</p>
          </motion.div>
        ) : (
          <motion.div key="idle" initial={{ opacity: 0 }} animate={{ opacity: 1 }} className="flex flex-col items-center">
            <div className="w-16 h-16 rounded-2xl bg-gradient-to-tr from-slate-800 to-slate-700 flex items-center justify-center shadow-xl border border-slate-600 mb-6">
              <UploadCloud className="w-8 h-8 text-blue-400" />
            </div>
            <h3 className="text-xl font-semibold text-white mb-2">Initialize Data Uplink</h3>
            <p className="text-slate-400 text-sm max-w-sm text-center">
              Drag & Drop the .docx performance sheet here.
            </p>
          </motion.div>
        )}
      </AnimatePresence>
    </motion.div>
  );
}
