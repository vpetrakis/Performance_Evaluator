"use client";
import { useState } from "react";
import { motion } from "framer-motion";
import DragDropZone from "../components/DragDropZone";
import { AlertTriangle, CheckCircle, Activity } from "lucide-react";

export default function Home() {
  const [report, setReport] = useState<any>(null);

  const handleFileUpload = async (file: File) => {
    const formData = new FormData();
    formData.append("file", file);

    try {
      const res = await fetch("http://localhost:8000/api/evaluate", {
        method: "POST",
        body: formData,
      });
      const data = await res.json();
      setTimeout(() => setReport(data), 800); // Artificial delay for premium feel
    } catch (error) {
      console.error("Evaluation failed", error);
    }
  };

  return (
    <main className="min-h-screen p-10 font-sans selection:bg-blue-500/30">
      <div className="max-w-5xl mx-auto">
        <header className="mb-12 text-center">
          <h1 className="text-4xl font-bold bg-clip-text text-transparent bg-gradient-to-r from-blue-400 to-emerald-400 mb-2">
            Performance Evaluator
          </h1>
          <p className="text-slate-400 tracking-wide uppercase text-sm font-semibold">Stage 1 Diagnostics</p>
        </header>

        {!report ? (
          <DragDropZone onUpload={handleFileUpload} />
        ) : (
          <motion.div initial={{ opacity: 0, y: 20 }} animate={{ opacity: 1, y: 0 }} className="space-y-6">
            
            {/* Executive Status Banner */}
            <div className={`p-6 rounded-2xl flex items-center gap-4 border ${
              report.status === "RED" ? "bg-red-950/30 border-red-500/50 text-red-400" :
              report.status === "YELLOW" ? "bg-amber-950/30 border-amber-500/50 text-amber-400" :
              "bg-emerald-950/30 border-emerald-500/50 text-emerald-400"
            }`}>
              {report.status === "GREEN" ? <CheckCircle className="w-8 h-8" /> : <AlertTriangle className="w-8 h-8" />}
              <div>
                <h2 className="text-2xl font-bold">System Status: {report.status}</h2>
                <p className="opacity-80 font-medium">
                  {report.snapshot.vessel_name} Engine operating at {report.snapshot.rpm} RPM
                </p>
              </div>
            </div>

            {/* Alerts Matrix */}
            {report.alerts.length > 0 && (
              <div className="bg-slate-800/50 border border-slate-700 rounded-2xl p-6">
                <h3 className="text-lg font-semibold text-white mb-4 flex items-center gap-2">
                  <Activity className="w-5 h-5 text-blue-400" /> Active Anomalies Detected
                </h3>
                <ul className="space-y-3">
                  {report.alerts.map((alert: str, idx: int) => (
                    <motion.li 
                      initial={{ opacity: 0, x: -10 }} animate={{ opacity: 1, x: 0 }} transition={{ delay: idx * 0.1 }}
                      key={idx} className="flex items-start gap-3 text-slate-300 bg-slate-900/50 p-4 rounded-xl border border-slate-700/50"
                    >
                      <span className="w-2 h-2 mt-2 rounded-full bg-amber-500 shrink-0" />
                      {alert}
                    </motion.li>
                  ))}
                </ul>
              </div>
            )}
            
            <button 
              onClick={() => setReport(null)}
              className="mt-8 px-6 py-3 bg-slate-800 hover:bg-slate-700 text-white rounded-xl font-medium transition-colors border border-slate-600"
            >
              Evaluate New Dataset
            </button>
          </motion.div>
        )}
      </div>
    </main>
  );
}
