import { useState, useCallback } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";

function normalizeHeader(h) {
  return h
    .replace(/^\uFEFF/, "")
    .trim()
    .toLowerCase()
    .replace(/[\s_]+/g, " ");
}

function findCol(headers, ...keys) {
  for (const k of keys) {
    const idx = headers.findIndex((h) =>
      normalizeHeader(h).includes(k.toLowerCase()),
    );
    if (idx !== -1) return idx;
  }
  return -1;
}

function ScoreCell({ value, onCopy }) {
  const [showCopy, setShowCopy] = useState(false);
  const cls =
    value > 0
      ? "text-emerald-600 dark:text-emerald-400 font-medium"
      : value < 0
        ? "text-red-500 dark:text-red-400 font-medium"
        : "text-gray-400 dark:text-gray-500";

  const handleCopy = () => {
    navigator.clipboard.writeText(value.toFixed(2));
    if (onCopy) onCopy();
  };

  return (
    <div
      className="relative inline-flex items-center gap-1 group cursor-pointer"
      onMouseEnter={() => setShowCopy(true)}
      onMouseLeave={() => setShowCopy(false)}
      onClick={handleCopy}
      title="Click to copy"
    >
      <span className={cls}>{value.toFixed(2)}</span>
      {showCopy && (
        <svg
          xmlns="http://www.w3.org/2000/svg"
          className="w-3 h-3 text-gray-400 hover:text-gray-600 dark:hover:text-gray-300 transition-colors"
          fill="none"
          viewBox="0 0 24 24"
          stroke="currentColor"
          strokeWidth={2}
        >
          <path
            strokeLinecap="round"
            strokeLinejoin="round"
            d="M8 16H6a2 2 0 01-2-2V6a2 2 0 012-2h8a2 2 0 012 2v2m-6 12h8a2 2 0 002-2v-8a2 2 0 00-2-2h-8a2 2 0 00-2 2v8a2 2 0 002 2z"
          />
        </svg>
      )}
    </div>
  );
}

function FinalPointCell({ value, onCopy }) {
  const [showCopy, setShowCopy] = useState(false);

  const handleCopy = () => {
    navigator.clipboard.writeText(value.toFixed(2));
    if (onCopy) onCopy();
  };

  return (
    <div
      className="relative inline-flex items-center gap-1 group cursor-pointer"
      onMouseEnter={() => setShowCopy(true)}
      onMouseLeave={() => setShowCopy(false)}
      onClick={handleCopy}
      title="Click to copy"
    >
      <span
        className={`font-semibold text-xs ${value > 0 ? "text-blue-600 dark:text-blue-400" : value < 0 ? "text-red-500 dark:text-red-400" : "text-gray-400 dark:text-gray-500"}`}
      >
        {value.toFixed(2)}
      </span>
      {showCopy && (
        <svg
          xmlns="http://www.w3.org/2000/svg"
          className="w-3 h-3 text-gray-400 hover:text-gray-600 dark:hover:text-gray-300 transition-colors"
          fill="none"
          viewBox="0 0 24 24"
          stroke="currentColor"
          strokeWidth={2}
        >
          <path
            strokeLinecap="round"
            strokeLinejoin="round"
            d="M8 16H6a2 2 0 01-2-2V6a2 2 0 012-2h8a2 2 0 012 2v2m-6 12h8a2 2 0 002-2v-8a2 2 0 00-2-2h-8a2 2 0 00-2 2v8a2 2 0 002 2z"
          />
        </svg>
      )}
    </div>
  );
}

function TrackerBadge({ tracker }) {
  const t = (tracker || "").toLowerCase();
  let cls = "";
  if (t === "be") {
    cls = "bg-blue-100 text-blue-700 dark:bg-blue-900/30 dark:text-blue-300";
  } else if (t === "fe") {
    cls =
      "bg-green-100 text-green-700 dark:bg-green-900/30 dark:text-green-300";
  } else {
    cls =
      "bg-purple-100 text-purple-700 dark:bg-purple-900/30 dark:text-purple-300";
  }
  return (
    <span className={`px-2 py-0.5 rounded text-xs font-medium ${cls}`}>
      {tracker}
    </span>
  );
}

export default function KPIScoreCalculator() {
  const [rows, setRows] = useState([]);
  const [search, setSearch] = useState("");
  const [trackerFilter, setTrackerFilter] = useState("");
  const [scoreFilter, setScoreFilter] = useState("");
  const [dragging, setDragging] = useState(false);
  const [darkMode, setDarkMode] = useState(true);
  const [showSettings, setShowSettings] = useState(false);
  const [qualityPercent, setQualityPercent] = useState(5);
  const [onTimePercent, setOnTimePercent] = useState(5);
  const [editingQuality, setEditingQuality] = useState(null);
  const [editingQualityPercent, setEditingQualityPercent] = useState("");
  const [showDownloadMenu, setShowDownloadMenu] = useState(false);
  const [copyMessage, setCopyMessage] = useState("");

  function processRows(rawRows) {
    if (rawRows.length < 2) return;
    const headers = rawRows[0];
    const iId = findCol(headers, "#", "id", "issue");
    const iTracker = findCol(headers, "tracker");
    const iSubject = findCol(headers, "subject", "task", "title");
    const iAssignee = findCol(headers, "assignee");
    const iEst = findCol(headers, "estimated time", "estimated");
    const iSpent = findCol(headers, "total spent time", "spent time", "spent");

    const parsed = rawRows
      .slice(1)
      .map((r) => {
        const est = parseFloat(r[iEst]) || 0;
        const spent = parseFloat(r[iSpent]) || 0;
        const score = est - spent;

        const taskScore = score < 0 ? score + est : est - score;
        const onTimeScore = (onTimePercent / 100) * taskScore;
        const qualityScore = (qualityPercent / 100) * taskScore;
        const finalPoint = taskScore + onTimeScore + qualityScore;
        return {
          id: iId >= 0 ? r[iId] : "",
          tracker: iTracker >= 0 ? r[iTracker] : "",
          subject: iSubject >= 0 ? r[iSubject] : "",
          assignee: iAssignee >= 0 ? r[iAssignee] : "",
          est,
          spent,
          onTime: "yes",
          taskScore,
          onTimeScore,
          qualityScore,
          qualityPercentValue: qualityPercent,
          finalPoint,
        };
      })
      .filter((r) => r.est > 0 || r.spent > 0 || r.subject);

    setRows(parsed);
    setSearch("");
    setTrackerFilter("");
    setScoreFilter("");
  }

  function handleCSV(text) {
    const result = Papa.parse(text, { skipEmptyLines: true });
    processRows(result.data);
  }

  function handleXLSX(buffer) {
    const wb = XLSX.read(buffer, { type: "array" });
    const ws = wb.Sheets[wb.SheetNames[0]];
    handleCSV(XLSX.utils.sheet_to_csv(ws));
  }

  function handleFile(file) {
    if (!file) return;
    const ext = file.name.split(".").pop().toLowerCase();
    if (ext === "csv") {
      const reader = new FileReader();
      reader.onload = (e) => handleCSV(e.target.result);
      reader.readAsText(file);
    } else {
      const reader = new FileReader();
      reader.onload = (e) => handleXLSX(e.target.result);
      reader.readAsArrayBuffer(file);
    }
  }

  const onInputChange = (e) => handleFile(e.target.files[0]);

  const onDrop = useCallback((e) => {
    e.preventDefault();
    setDragging(false);
    handleFile(e.dataTransfer.files[0]);
  }, []);

  const handleCopySuccess = () => {
    setCopyMessage("Copied!");
    setTimeout(() => setCopyMessage(""), 1500);
  };

  const handleOnTimeChange = (index, value) => {
    const updatedRows = [...rows];
    updatedRows[index].onTime = value;

    if (value === "no") {
      updatedRows[index].onTimeScore = 0;
    } else {
      updatedRows[index].onTimeScore =
        (onTimePercent / 100) * updatedRows[index].taskScore;
    }

    updatedRows[index].finalPoint =
      updatedRows[index].taskScore +
      updatedRows[index].onTimeScore +
      updatedRows[index].qualityScore;

    setRows(updatedRows);
  };

  const handleQualityEdit = (index, percentValue) => {
    const num = parseFloat(percentValue);
    if (!isNaN(num) && num <= 100) {
      const updatedRows = [...rows];
      const newQualityScore = (num / 100) * updatedRows[index].taskScore;
      updatedRows[index].qualityScore = newQualityScore;
      updatedRows[index].qualityPercentValue = num;
      updatedRows[index].finalPoint =
        updatedRows[index].taskScore +
        updatedRows[index].onTimeScore +
        newQualityScore;
      setRows(updatedRows);
    }
    setEditingQuality(null);
    setEditingQualityPercent("");
  };

  const handleQualityPercentChange = (value) => {
    const num = parseFloat(value);
    if (!isNaN(num) && num <= 100) {
      setQualityPercent(num);
      if (rows.length > 0) {
        const updatedRows = rows.map((r) => {
          const qualityScore = (num / 100) * r.taskScore;
          const finalPoint = r.taskScore + r.onTimeScore + qualityScore;
          return {
            ...r,
            qualityScore,
            qualityPercentValue: num,
            finalPoint,
          };
        });
        setRows(updatedRows);
      }
    }
  };

  const handleOnTimePercentChange = (value) => {
    const num = parseFloat(value);
    if (!isNaN(num) && num <= 100) {
      setOnTimePercent(num);
      if (rows.length > 0) {
        const updatedRows = rows.map((r) => {
          let onTimeScore = r.onTimeScore;
          if (r.onTime === "yes") {
            onTimeScore = (num / 100) * r.taskScore;
          }
          const finalPoint = r.taskScore + onTimeScore + r.qualityScore;
          return {
            ...r,
            onTimeScore,
            finalPoint,
          };
        });
        setRows(updatedRows);
      }
    }
  };

  const downloadCSV = () => {
    const headers = [
      "#",
      "Tracker",
      "Subject",
      "Assignee",
      "Est (h)",
      "On Time",
      "Spent (h)",
      "Task Score",
      "On-Time Score",
      "Quality Score",
      "Final Point",
    ];
    const csvData = filtered.map((r) => [
      r.id,
      r.tracker,
      r.subject,
      r.assignee,
      r.est.toFixed(2),
      r.onTime,
      r.spent.toFixed(2),
      r.taskScore.toFixed(2),
      r.onTimeScore.toFixed(2),
      r.qualityScore.toFixed(2),
      r.finalPoint.toFixed(2),
    ]);
    const csvContent = [headers, ...csvData]
      .map((row) => row.join(","))
      .join("\n");
    const blob = new Blob([csvContent], { type: "text/csv" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `kpi_scores_${new Date().toISOString().slice(0, 19)}.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    setShowDownloadMenu(false);
  };

  const downloadPDF = () => {
    const doc = new jsPDF({ orientation: "landscape" });
    doc.text("KPI Score Report", 14, 10);
    doc.text(`Generated: ${new Date().toLocaleString()}`, 14, 18);

    const tableData = filtered.map((r) => [
      r.id,
      r.tracker,
      r.subject.length > 30 ? r.subject.substring(0, 27) + "..." : r.subject,
      r.assignee,
      r.est.toFixed(2),
      r.onTime,
      r.spent.toFixed(2),
      r.taskScore.toFixed(2),
      r.onTimeScore.toFixed(2),
      r.qualityScore.toFixed(2),
      r.finalPoint.toFixed(2),
    ]);

    autoTable(doc, {
      head: [
        [
          "#",
          "Tracker",
          "Subject",
          "Assignee",
          "Est",
          "On Time",
          "Spent",
          "Task Score",
          "On-Time Score",
          "Quality Score",
          "Final Point",
        ],
      ],
      body: tableData,
      startY: 22,
      theme: "striped",
      styles: { fontSize: 8, cellPadding: 2 },
      headStyles: {
        fillColor: [59, 130, 246],
        textColor: 255,
        fontSize: 8,
        fontStyle: "bold",
      },
      alternateRowStyles: { fillColor: [241, 245, 249] },
    });

    doc.save(`kpi_scores_${new Date().toISOString().slice(0, 19)}.pdf`);
    setShowDownloadMenu(false);
  };

  const trackers = [...new Set(rows.map((r) => r.tracker).filter(Boolean))];

  const filtered = rows.filter((r) => {
    const q = search.toLowerCase();
    if (
      q &&
      !r.subject.toLowerCase().includes(q) &&
      !r.assignee.toLowerCase().includes(q)
    )
      return false;
    if (trackerFilter && r.tracker !== trackerFilter) return false;
    if (scoreFilter === "pos" && r.taskScore < 0) return false;
    if (scoreFilter === "neg" && r.taskScore >= 0) return false;
    return true;
  });

  const total = rows.length;
  const totalFinalPoint = filtered.reduce((sum, r) => sum + r.finalPoint, 0);
  const avgTask = total ? rows.reduce((s, r) => s + r.taskScore, 0) / total : 0;
  const avgOnTime = total
    ? rows.reduce((s, r) => s + r.onTimeScore, 0) / total
    : 0;
  const avgQuality = total
    ? rows.reduce((s, r) => s + r.qualityScore, 0) / total
    : 0;
  const avgFinal = total
    ? rows.reduce((s, r) => s + r.finalPoint, 0) / total
    : 0;
  const positiveCount = rows.filter((r) => r.taskScore >= 0).length;

  return (
    <div
      className={`min-h-screen font-sans transition-colors duration-200 ${darkMode ? "dark bg-gray-900" : "bg-white"}`}
    >
      <div className="max-w-7xl mx-auto px-4 py-6">
        {/* Header */}
        <div className="flex items-start justify-between mb-5 flex-wrap gap-3">
          <div>
            <h1
              className={`text-xl font-medium ${darkMode ? "text-gray-100" : "text-gray-900"}`}
            >
              KPI Score Calculator
            </h1>
            <p
              className={`text-xs mt-0.5 ${darkMode ? "text-gray-400" : "text-gray-500"}`}
            >
              Upload your task sheet to calculate performance scores
            </p>
          </div>
          <div className="flex gap-2">
            {rows.length > 0 && (
              <div className="relative">
                <button
                  onClick={() => setShowDownloadMenu(!showDownloadMenu)}
                  className="flex items-center gap-1.5 px-3 py-1.5 border rounded-lg text-xs transition-colors"
                  style={{
                    borderColor: darkMode ? "#374151" : "#e5e7eb",
                    backgroundColor: darkMode ? "#1f2937" : "white",
                    color: darkMode ? "#d1d5db" : "#374151",
                  }}
                >
                  <svg
                    xmlns="http://www.w3.org/2000/svg"
                    className="w-3.5 h-3.5"
                    fill="none"
                    viewBox="0 0 24 24"
                    stroke="currentColor"
                    strokeWidth={1.5}
                  >
                    <path
                      strokeLinecap="round"
                      strokeLinejoin="round"
                      d="M3 16.5v2.25A2.25 2.25 0 005.25 21h13.5A2.25 2.25 0 0021 18.75V16.5M16.5 12L12 16.5m0 0L7.5 12m4.5 4.5V3"
                    />
                  </svg>
                  Download
                  <svg
                    xmlns="http://www.w3.org/2000/svg"
                    className="w-3 h-3"
                    fill="none"
                    viewBox="0 0 24 24"
                    stroke="currentColor"
                    strokeWidth={2}
                  >
                    <path
                      strokeLinecap="round"
                      strokeLinejoin="round"
                      d="M19 9l-7 7-7-7"
                    />
                  </svg>
                </button>
                {showDownloadMenu && (
                  <div
                    className={`absolute right-0 mt-1 w-28 rounded-lg shadow-lg border z-10 ${darkMode ? "bg-gray-800 border-gray-700" : "bg-white border-gray-200"}`}
                  >
                    <button
                      onClick={downloadCSV}
                      className={`block w-full text-left px-3 py-2 text-xs rounded-t-lg ${darkMode ? "hover:bg-gray-700 text-gray-200" : "hover:bg-gray-50 text-gray-700"}`}
                    >
                      📄 CSV
                    </button>
                    <button
                      onClick={downloadPDF}
                      className={`block w-full text-left px-3 py-2 text-xs rounded-b-lg ${darkMode ? "hover:bg-gray-700 text-gray-200" : "hover:bg-gray-50 text-gray-700"}`}
                    >
                      📑 PDF
                    </button>
                  </div>
                )}
              </div>
            )}
            <button
              onClick={() => setShowSettings(!showSettings)}
              className="flex items-center gap-1.5 px-3 py-1.5 border rounded-lg text-xs transition-colors"
              style={{
                borderColor: darkMode ? "#374151" : "#e5e7eb",
                backgroundColor: darkMode ? "#1f2937" : "white",
                color: darkMode ? "#d1d5db" : "#374151",
              }}
            >
              <svg
                xmlns="http://www.w3.org/2000/svg"
                className="w-3.5 h-3.5"
                fill="none"
                viewBox="0 0 24 24"
                stroke="currentColor"
                strokeWidth={1.5}
              >
                <path
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  d="M10.325 4.317c.426-1.756 2.924-1.756 3.35 0a1.724 1.724 0 002.573 1.066c1.543-.94 3.31.826 2.37 2.37a1.724 1.724 0 001.065 2.572c1.756.426 1.756 2.924 0 3.35a1.724 1.724 0 00-1.066 2.573c.94 1.543-.826 3.31-2.37 2.37a1.724 1.724 0 00-2.572 1.065c-.426 1.756-2.924 1.756-3.35 0a1.724 1.724 0 00-2.573-1.066c-1.543.94-3.31-.826-2.37-2.37a1.724 1.724 0 00-1.065-2.572c-1.756-.426-1.756-2.924 0-3.35a1.724 1.724 0 001.066-2.573c-.94-1.543.826-3.31 2.37-2.37.996.608 2.296.07 2.572-1.065z"
                />
                <path
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  d="M15 12a3 3 0 11-6 0 3 3 0 016 0z"
                />
              </svg>
              Settings
            </button>
            <button
              onClick={() => setDarkMode(!darkMode)}
              className="flex items-center gap-1.5 px-3 py-1.5 border rounded-lg text-xs transition-colors"
              style={{
                borderColor: darkMode ? "#374151" : "#e5e7eb",
                backgroundColor: darkMode ? "#1f2937" : "white",
                color: darkMode ? "#d1d5db" : "#374151",
              }}
            >
              {darkMode ? (
                <svg
                  xmlns="http://www.w3.org/2000/svg"
                  className="w-3.5 h-3.5"
                  fill="none"
                  viewBox="0 0 24 24"
                  stroke="currentColor"
                  strokeWidth={1.5}
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    d="M12 3v1m0 16v1m9-9h-1M4 12H3m15.364 6.364l-.707-.707M6.343 6.343l-.707-.707m12.728 0l-.707.707M6.343 17.657l-.707.707M16 12a4 4 0 11-8 0 4 4 0 018 0z"
                  />
                </svg>
              ) : (
                <svg
                  xmlns="http://www.w3.org/2000/svg"
                  className="w-3.5 h-3.5"
                  fill="none"
                  viewBox="0 0 24 24"
                  stroke="currentColor"
                  strokeWidth={1.5}
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    d="M20.354 15.354A9 9 0 018.646 3.646 9.003 9.003 0 0012 21a9.003 9.003 0 008.354-5.646z"
                  />
                </svg>
              )}
              {darkMode ? "Light" : "Dark"}
            </button>
            <label
              className="cursor-pointer flex items-center gap-1.5 px-3 py-1.5 border rounded-lg text-xs transition-colors"
              style={{
                borderColor: darkMode ? "#374151" : "#e5e7eb",
                backgroundColor: darkMode ? "#1f2937" : "white",
                color: darkMode ? "#d1d5db" : "#374151",
              }}
            >
              <svg
                xmlns="http://www.w3.org/2000/svg"
                className="w-3.5 h-3.5"
                fill="none"
                viewBox="0 0 24 24"
                stroke="currentColor"
                strokeWidth={1.5}
              >
                <path
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  d="M3 16.5v2.25A2.25 2.25 0 005.25 21h13.5A2.25 2.25 0 0021 18.75V16.5m-13.5-9L12 3m0 0l4.5 4.5M12 3v13.5"
                />
              </svg>
              Upload
              <input
                type="file"
                accept=".csv,.xlsx,.xls"
                className="hidden"
                onChange={onInputChange}
              />
            </label>
          </div>
        </div>

        {/* Copy Success Message */}
        {copyMessage && (
          <div className="fixed bottom-4 right-4 z-50 bg-green-500 text-white text-xs px-3 py-2 rounded-lg shadow-lg animate-pulse">
            {copyMessage}
          </div>
        )}

        {/* Settings Panel */}
        {showSettings && rows.length > 0 && (
          <div
            className={`mb-5 p-3 rounded-lg border ${darkMode ? "bg-gray-800 border-gray-700" : "bg-gray-50 border-gray-200"}`}
          >
            <div className="flex items-center justify-between mb-2">
              <h3
                className={`text-xs font-medium ${darkMode ? "text-gray-200" : "text-gray-700"}`}
              >
                Score Calculation Settings
              </h3>
              <button
                onClick={() => setShowSettings(false)}
                className={`text-xs ${darkMode ? "text-gray-400 hover:text-gray-300" : "text-gray-400 hover:text-gray-600"}`}
              >
                Close
              </button>
            </div>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
              <div>
                <label
                  className={`block text-xs mb-0.5 ${darkMode ? "text-gray-400" : "text-gray-500"}`}
                >
                  On-Time Score (%)
                </label>
                <div className="flex items-center gap-2">
                  <input
                    type="range"
                    min="-100"
                    max="100"
                    step="1"
                    value={onTimePercent}
                    onChange={(e) => handleOnTimePercentChange(e.target.value)}
                    className="flex-1 h-1.5 rounded-lg appearance-none cursor-pointer bg-gray-200 dark:bg-gray-700"
                  />
                  <input
                    type="text"
                    min="-100"
                    max="100"
                    step="1"
                    value={onTimePercent}
                    onChange={(e) => handleOnTimePercentChange(e.target.value)}
                    className={`w-14 text-center text-xs px-1.5 py-0.5 border rounded-lg ${darkMode ? "bg-gray-700 border-gray-600 text-gray-200" : "bg-white border-gray-200 text-gray-700"}`}
                  />
                  <span
                    className={`text-xs ${darkMode ? "text-gray-400" : "text-gray-500"}`}
                  >
                    %
                  </span>
                </div>
              </div>
              <div>
                <label
                  className={`block text-xs mb-0.5 ${darkMode ? "text-gray-400" : "text-gray-500"}`}
                >
                  Quality Score (%)
                </label>
                <div className="flex items-center gap-2">
                  <input
                    type="range"
                    min="-100"
                    max="100"
                    step="1"
                    value={qualityPercent}
                    onChange={(e) => handleQualityPercentChange(e.target.value)}
                    className="flex-1 h-1.5 rounded-lg appearance-none cursor-pointer bg-gray-200 dark:bg-gray-700"
                  />
                  <input
                    type="text"
                    min="-100"
                    max="100"
                    step="1"
                    value={qualityPercent}
                    onChange={(e) => handleQualityPercentChange(e.target.value)}
                    className={`w-14 text-center text-xs px-1.5 py-0.5 border rounded-lg ${darkMode ? "bg-gray-700 border-gray-600 text-gray-200" : "bg-white border-gray-200 text-gray-700"}`}
                  />
                  <span
                    className={`text-xs ${darkMode ? "text-gray-400" : "text-gray-500"}`}
                  >
                    %
                  </span>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* Upload area */}
        {rows.length === 0 && (
          <div
            onDragOver={(e) => {
              e.preventDefault();
              setDragging(true);
            }}
            onDragLeave={() => setDragging(false)}
            onDrop={onDrop}
            onClick={() => document.getElementById("hiddenFileInput").click()}
            className={`border-2 border-dashed rounded-xl p-12 text-center cursor-pointer transition-colors mb-5
              ${dragging ? "border-blue-400 bg-blue-50 dark:bg-blue-900/20 dark:border-blue-500" : darkMode ? "border-gray-700 bg-gray-800 hover:bg-gray-750" : "border-gray-200 bg-gray-50 hover:bg-gray-100"}`}
          >
            <svg
              xmlns="http://www.w3.org/2000/svg"
              className={`w-8 h-8 mx-auto mb-2 ${darkMode ? "text-gray-500" : "text-gray-400"}`}
              fill="none"
              viewBox="0 0 24 24"
              stroke="currentColor"
              strokeWidth={1.2}
            >
              <path
                strokeLinecap="round"
                strokeLinejoin="round"
                d="M19.5 14.25v-2.625a3.375 3.375 0 00-3.375-3.375h-1.5A1.125 1.125 0 0113.5 7.125v-1.5a3.375 3.375 0 00-3.375-3.375H8.25m0 12.75h7.5m-7.5 3H12M10.5 2.25H5.625c-.621 0-1.125.504-1.125 1.125v17.25c0 .621.504 1.125 1.125 1.125h12.75c.621 0 1.125-.504 1.125-1.125V11.25a9 9 0 00-9-9z"
              />
            </svg>
            <p
              className={`text-xs mb-0.5 ${darkMode ? "text-gray-400" : "text-gray-600"}`}
            >
              Click or drag & drop your file here
            </p>
            <p
              className={`text-[11px] ${darkMode ? "text-gray-500" : "text-gray-400"}`}
            >
              Supports .csv, .xlsx, .xls — needs Estimated time & Total spent
              time columns
            </p>
            <input
              id="hiddenFileInput"
              type="file"
              accept=".csv,.xlsx,.xls"
              className="hidden"
              onChange={onInputChange}
            />
          </div>
        )}

        {/* Dashboard */}
        {rows.length > 0 && (
          <>
            {/* Filters */}
            <div className="flex flex-wrap gap-2 mb-3">
              <input
                type="text"
                placeholder="Search..."
                value={search}
                onChange={(e) => setSearch(e.target.value)}
                className={`flex-1 min-w-40 text-xs px-2 py-1.5 border rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-100 dark:focus:ring-blue-900/50 ${darkMode ? "bg-gray-800 border-gray-700 text-gray-200 placeholder-gray-500 focus:border-blue-600" : "bg-white border-gray-200 text-gray-800 placeholder-gray-400 focus:border-blue-300"}`}
              />
              <select
                value={trackerFilter}
                onChange={(e) => setTrackerFilter(e.target.value)}
                className={`text-xs px-2 py-1.5 border rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-100 dark:focus:ring-blue-900/50 ${darkMode ? "bg-gray-800 border-gray-700 text-gray-200" : "bg-white border-gray-200 text-gray-700"}`}
              >
                <option value="">All trackers</option>
                {trackers.map((t) => (
                  <option key={t} value={t}>
                    {t}
                  </option>
                ))}
              </select>
              <select
                value={scoreFilter}
                onChange={(e) => setScoreFilter(e.target.value)}
                className={`text-xs px-2 py-1.5 border rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-100 dark:focus:ring-blue-900/50 ${darkMode ? "bg-gray-800 border-gray-700 text-gray-200" : "bg-white border-gray-200 text-gray-700"}`}
              >
                <option value="">All scores</option>
                <option value="pos">Positive</option>
                <option value="neg">Negative</option>
              </select>
              <label
                className={`cursor-pointer flex items-center gap-1.5 px-2 py-1.5 border rounded-lg text-xs transition-colors ${darkMode ? "border-gray-700 text-gray-300 hover:bg-gray-800" : "border-gray-200 text-gray-600 hover:bg-gray-50"}`}
              >
                <svg
                  xmlns="http://www.w3.org/2000/svg"
                  className="w-3.5 h-3.5"
                  fill="none"
                  viewBox="0 0 24 24"
                  stroke="currentColor"
                  strokeWidth={1.5}
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    d="M3 16.5v2.25A2.25 2.25 0 005.25 21h13.5A2.25 2.25 0 0021 18.75V16.5m-13.5-9L12 3m0 0l4.5 4.5M12 3v13.5"
                  />
                </svg>
                New
                <input
                  type="file"
                  accept=".csv,.xlsx,.xls"
                  className="hidden"
                  onChange={onInputChange}
                />
              </label>
            </div>

            {/* Table */}
            <div
              className={`overflow-x-auto rounded-lg border ${darkMode ? "border-gray-700" : "border-gray-200"}`}
            >
              <table
                className="w-full text-xs border-collapse"
                style={{ minWidth: 980 }}
              >
                <thead>
                  <tr
                    className={`${darkMode ? "bg-gray-800/50 text-gray-400" : "bg-gray-50 text-gray-500"} text-[11px] font-medium`}
                  >
                    <th className="px-2 py-2 text-left w-12">#</th>
                    <th className="px-2 py-2 text-left w-14">Tracker</th>
                    <th
                      className="px-2 py-2 text-left"
                      style={{ minWidth: 180 }}
                    >
                      Subject
                    </th>
                    <th className="px-2 py-2 text-left w-24">Assignee</th>
                    <th className="px-2 py-2 text-right w-14">Est (h)</th>
                    <th className="px-2 py-2 text-center w-20">On Time</th>
                    <th className="px-2 py-2 text-right w-14">Spent (h)</th>
                    <th className="px-2 py-2 text-right w-20">Task Score</th>
                    <th className="px-2 py-2 text-right w-24">On-Time Score</th>
                    <th className="px-2 py-2 text-right w-20">Quality Score</th>
                    <th
                      className={`px-2 py-2 text-right w-20 ${darkMode ? "text-blue-400" : "text-blue-600"}`}
                    >
                      Final Point
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {filtered.length === 0 ? (
                    <tr>
                      <td
                        colSpan={11}
                        className={`text-center py-8 text-xs ${darkMode ? "text-gray-500" : "text-gray-400"}`}
                      >
                        No tasks match your filters
                      </td>
                    </tr>
                  ) : (
                    filtered.map((r, i) => (
                      <tr
                        key={i}
                        className={`border-t ${darkMode ? "border-gray-800 hover:bg-gray-800/50" : "border-gray-100 hover:bg-gray-50"} transition-colors`}
                      >
                        <td
                          className={`px-2 py-2 text-[11px] ${darkMode ? "text-gray-500" : "text-gray-400"}`}
                        >
                          {r.id}
                        </td>
                        <td className="px-2 py-2">
                          <TrackerBadge tracker={r.tracker} />
                        </td>
                        <td
                          className={`px-2 py-2 max-w-xs truncate ${darkMode ? "text-gray-200" : "text-gray-800"}`}
                          title={r.subject}
                        >
                          {r.subject}
                        </td>
                        <td
                          className={`px-2 py-1 text-[11px] truncate ${darkMode ? "text-gray-400" : "text-gray-500"}`}
                        >
                          {r.assignee}
                        </td>
                        <td
                          className={`px-2 py-2 text-right ${darkMode ? "text-gray-300" : "text-gray-700"}`}
                        >
                          {r.est.toFixed(2)}
                        </td>
                        <td className="px-2 py-2 text-center">
                          <select
                            value={r.onTime}
                            onChange={(e) =>
                              handleOnTimeChange(i, e.target.value)
                            }
                            className={`text-[11px] px-1 py-0.5 border rounded ${darkMode ? "bg-gray-700 border-gray-600 text-gray-200" : "bg-white border-gray-200 text-gray-700"}`}
                          >
                            <option value="yes">Yes</option>
                            <option value="no">No</option>
                          </select>
                        </td>
                        <td
                          className={`px-2 py-2 text-right ${darkMode ? "text-gray-300" : "text-gray-700"}`}
                        >
                          {r.spent.toFixed(2)}
                        </td>
                        <td className="px-2 py-2 text-right">
                          <ScoreCell
                            value={r.taskScore}
                            onCopy={handleCopySuccess}
                          />
                        </td>
                        <td className="px-2 py-2 text-right">
                          <ScoreCell
                            value={r.onTimeScore}
                            onCopy={handleCopySuccess}
                          />
                        </td>
                        <td className="px-2 py-2 text-right">
                          {editingQuality === i ? (
                            <div className="flex items-center gap-1 justify-end">
                              <input
                                type="number"
                                step="1"
                                min="0"
                                max="100"
                                value={editingQualityPercent}
                                onChange={(e) =>
                                  setEditingQualityPercent(e.target.value)
                                }
                                onBlur={() =>
                                  handleQualityEdit(i, editingQualityPercent)
                                }
                                onKeyPress={(e) => {
                                  if (e.key === "Enter") {
                                    handleQualityEdit(i, editingQualityPercent);
                                  }
                                }}
                                placeholder="%"
                                autoFocus
                                className={`w-14 text-right text-xs px-1 py-0.5 border rounded ${darkMode ? "bg-gray-700 border-gray-600 text-gray-200" : "bg-white border-gray-200 text-gray-700"}`}
                              />
                              <span className="text-[10px]">%</span>
                            </div>
                          ) : (
                            <div className="flex items-center justify-end gap-1">
                              <span
                                onClick={() => {
                                  setEditingQuality(i);
                                  setEditingQualityPercent(
                                    r.qualityPercentValue?.toString() ||
                                      qualityPercent.toString(),
                                  );
                                }}
                                className="cursor-pointer hover:bg-gray-100 dark:hover:bg-gray-700 px-1 py-0.5 rounded inline-flex items-center gap-1"
                              >
                                <ScoreCell
                                  value={r.qualityScore}
                                  onCopy={handleCopySuccess}
                                />
                                <span className="text-[10px] text-gray-400">
                                  ({r.qualityPercentValue || qualityPercent}%)
                                </span>
                              </span>
                            </div>
                          )}
                        </td>
                        <td className="px-2 py-2 text-right">
                          <FinalPointCell
                            value={r.finalPoint}
                            onCopy={handleCopySuccess}
                          />
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>

            <div className="flex justify-between items-center mt-2">
              <p
                className={`text-[11px] ${darkMode ? "text-gray-500" : "text-gray-400"}`}
              >
                Showing {filtered.length} of {total} tasks
              </p>
              <div
                className={`text-right text-xs font-medium px-3 py-1 rounded-lg ${darkMode ? "bg-gray-800 text-blue-400" : "bg-blue-50 text-blue-700"}`}
              >
                Total Final Point: {totalFinalPoint.toFixed(2)}
              </div>
            </div>
          </>
        )}
      </div>
    </div>
  );
}
