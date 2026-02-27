import { useState, useRef, useCallback } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";

const OUTPUT_COLUMNS = [
  "Stock #",
  "Availability",
  "Shape",
  "Weight",
  "Color",
  "Clarity",
  "Cut Grade",
  "Polish",
  "Country of Polishing",
  "Symmetry",
  "Fluorescence Intensity",
  "Fluorescence Color",
  "Measurements",
  "Lab",
  "Report #",
  "RapNet Price",
  "Fancy Color",
  "Fancy Color Intensity",
  "Fancy Color Overtone",
  "Depth %",
  "Table %",
  "Girdle Thin",
  "Girdle Thick",
  "Girdle %",
  "Girdle Condition",
  "Culet Condition",
  "Crown Height",
  "Crown Angle",
  "Pavilion Depth",
  "Pavilion Angle",
  "City",
  "Country",
  "DiamondImage",
  "VideoURL",
  "Heart image",
  "Arrow Image",
  "Aset Image",
  "Milky",
  "Eye Clean",
  "Shade",
  "Brand",
  "Origin ",
  "Treatment",
  "Key to symbols",
  "Report Comments",
];

const DIRECT_MAP = {
  "Stock #": "Stock #",
  Availability: "Availability",
  Shape: "Shape",
  Weight: "Weight",
  Color: "Color",
  Clarity: "Clarity",
  "Cut Grade": "Cut Grade",
  Polish: "Polish",
  "Country of Polishing": "Country",
  Symmetry: "Symmetry",
  "Fluorescence Intensity": "Fluorescence Intensity",
  "Fluorescence Color": "Fluorescence Color",
  Measurements: "Measurements",
  Lab: "Lab",
"DiamondImage": "DiamondImage",
"Heart image": "Heart image",
"Arrow Image": "Arrow Image",
"Aset Image": "Aset Image",
"Origin ": "Origin ",
  "Fancy Color": "Fancy Color",
  "Fancy Color Intensity": "Fancy Color Intensity",
  "Fancy Color Overtone": "Fancy Color Overtone",
  "Depth %": "Depth %",
  "Table %": "Table %",
  "Girdle Thin": "Girdle Thin",
  "Girdle Thick": "Girdle Thick",
  "Girdle %": "Girdle %",
  "Girdle Condition": "Girdle Condition",
  "Culet Condition": "Culet Condition",
  "Crown Height": "Crown Height",
  "Crown Angle": "Crown Angle",
  "Pavilion Depth": "Pavilion Depth",
  "Pavilion Angle": "Pavilion Angle",
  City: "City",
  Country: "Country",
  Milky: "Milky",
  "Eye Clean": "Eye Clean",
  Shade: "Shade",
  Brand: "Brand",
  Treatment: "Treatment",
  "Key to symbols": "Key to symbols",
  "Report Comments": "Cert comment",
};



// +0e format fixture
function toIntString(val) {
  if (val === null || val === undefined || val === "") return "";
  const str = String(val).trim();
  if (!str) return "";

  const num = parseFloat(str);
  if (isNaN(num)) return str;

  return Math.round(num).toString();
}

function extractFirstUrl(cellVal) {
  if (!cellVal) return "";
  const str = String(cellVal).trim();
  if (!str) return "";

  const first = str.search(/https?:\/\//);
  if (first === -1) return "";

  const rest = str.slice(first + 8);
  const secondOffset = rest.search(/https?:\/\//);

  let raw;
  if (secondOffset === -1) {
    raw = str.slice(first).trim();
  } else {
    raw = str.slice(first, first + 8 + secondOffset).trim();
  }

  const wsMatch = raw.match(/^[^\s]+/);
  return wsMatch ? wsMatch[0] : "";
}

function pickMediaUrl(row) {
  const videoKeys = [
    "VideoURL 1",
    "VideoURL 2",
    "VideoURL 3",
    "VideoURL 4",
    "VideoURL 5",
  ];
  for (const key of videoKeys) {
    const url = extractFirstUrl(row[key]);
    if (url) return url;
  }
  const imageKeys = ["Image 1", "Image 2", "Image 3", "Image 4", "Image 5"];
  for (const key of imageKeys) {
    const url = extractFirstUrl(row[key]);
    if (url) return url;
  }
  return "";
}

function transformRow(row) {
  const out = {};
  for (const col of OUTPUT_COLUMNS) {
    if (col === "Report #") {
      out[col] = toIntString(row["Report #"]);
    } else if (col === "RapNet Price") {
      const v = row["RapNet Price"] ?? row["BuyNow $/ct"] ?? "";
      out[col] = toIntString(v);
    } else if (col === "VideoURL") {
      out[col] = pickMediaUrl(row);
    } else if (DIRECT_MAP[col]) {
      const raw = row[DIRECT_MAP[col]];
      out[col] =
        raw === null || raw === undefined
          ? ""
          : String(raw).trim() === "nan"
            ? ""
            : String(raw).trim();
    } else {
      out[col] = "";
    }
  }
  return out;
}

function generateXlsx(rows) {
  const ws = XLSX.utils.json_to_sheet(rows, { header: OUTPUT_COLUMNS });

  const range = XLSX.utils.decode_range(ws["!ref"]);
  const reportColIdx = OUTPUT_COLUMNS.indexOf("Report #");
  const rapnetColIdx = OUTPUT_COLUMNS.indexOf("RapNet Price");

  for (let R = range.s.r + 1; R <= range.e.r; R++) {
    [reportColIdx, rapnetColIdx].forEach((C) => {
      if (C < 0) return;
      const cellAddr = XLSX.utils.encode_cell({ r: R, c: C });
      const cell = ws[cellAddr];
      if (cell && cell.v !== "" && cell.v !== undefined) {
        const strVal = String(cell.v).trim();
        if (/^\d+$/.test(strVal)) {
          cell.t = "s";
          cell.v = strVal;
          cell.z = "@";
          delete cell.w;
        }
      }
    });
  }

  const videoColIdx = OUTPUT_COLUMNS.indexOf("VideoURL");
  if (videoColIdx >= 0) {
    for (let R = range.s.r + 1; R <= range.e.r; R++) {
      const cellAddr = XLSX.utils.encode_cell({ r: R, c: videoColIdx });
      const cell = ws[cellAddr];
      if (cell && cell.v && String(cell.v).startsWith("http")) {
        const url = String(cell.v).replace(/"/g, '""');
        cell.t = "s";
        cell.v = url;
        cell.f = `HYPERLINK("${url}","${url}")`;
        delete cell.w;
      }
    }
  }

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
  return XLSX.write(wb, { bookType: "xlsx", type: "array" });
}

// Ui design

export default function Home() {
  const [status, setStatus] = useState("idle");
  const [errorMsg, setErrorMsg] = useState("");
  const [isDragging, setIsDragging] = useState(false);
  const fileInputRef = useRef(null);
  const downloadLinkRef = useRef(null);

  const processFile = useCallback((file) => {
    if (!file) return;
    const name = file.name.toLowerCase();
    const isCSV = name.endsWith(".csv");
    const isXLSX = name.endsWith(".xlsx");
    const isXLS = name.endsWith(".xls");

    if (!isCSV && !isXLSX && !isXLS) {
      setStatus("error");
      setErrorMsg("Please upload a CSV, XLSX, or XLS file.");
      return;
    }

    setStatus("processing");
    setErrorMsg("");

    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        let rawData;

        if (isCSV) {
          const csvText = e.target.result;
          const result = Papa.parse(csvText, {
            header: true,
            skipEmptyLines: true,
            dynamicTyping: false,
          });
          rawData = result.data;
        } else {
          const wb = XLSX.read(e.target.result, { type: "array", raw: true });
          const ws = wb.Sheets[wb.SheetNames[0]];
          rawData = XLSX.utils.sheet_to_json(ws, { defval: "", raw: true });
        }

        const rows = rawData.map(transformRow);
        const xlsxBuffer = generateXlsx(rows);

        const blob = new Blob([xlsxBuffer], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        });
        const url = URL.createObjectURL(blob);

        setStatus("done");

        const link = document.createElement("a");
        link.href = url;
        link.download = "Output.xlsx";
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        setTimeout(() => URL.revokeObjectURL(url), 5000);
      } catch (err) {
        console.error(err);
        setStatus("error");
        setErrorMsg(err.message || "Something went wrong during conversion.");
      }
    };

    reader.onerror = () => {
      setStatus("error");
      setErrorMsg("Failed to read the file.");
    };

    if (isCSV) {
      reader.readAsText(file);
    } else {
      reader.readAsArrayBuffer(file);
    }
  }, []);

  const handleDrop = useCallback(
    (e) => {
      e.preventDefault();
      setIsDragging(false);
      const file = e.dataTransfer.files[0];
      processFile(file);
    },
    [processFile],
  );

  const handleFileChange = (e) => {
    processFile(e.target.files[0]);
    e.target.value = "";
  };

  return (
    <main className="min-h-screen bg-gradient-to-br from-slate-900 via-slate-800 to-slate-900 flex items-center justify-center p-6">
      <div className="w-full max-w-2xl">
        {/* Heading section */}
        <div className="text-center mb-10">
          <div className="inline-flex items-center justify-center w-16 h-16 rounded-2xl bg-blue-500/20 border border-blue-500/30 mb-4">
            <svg
              className="w-8 h-8 text-blue-400"
              fill="none"
              stroke="currentColor"
              viewBox="0 0 24 24"
            >
              <path
                strokeLinecap="round"
                strokeLinejoin="round"
                strokeWidth={2}
                d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M9 19l3 3m0 0l3-3m-3 3V10"
              />
            </svg>
          </div>
          <h1 className="text-3xl font-bold text-white mb-2">
            Rapnet - Nivoda
          </h1>
          <p className="text-slate-400 text-sm">
            Upload your CSV, XLSX, or XLS file and get a clean formatted XLSX
            instantly
          </p>
        </div>

        {/* File upload */}
        <div
          className={`relative rounded-2xl border-2 border-dashed transition-all duration-200 cursor-pointer
            ${isDragging ? "border-blue-400 bg-blue-500/10 scale-[1.01]" : "border-slate-600 bg-slate-800/50 hover:border-slate-500 hover:bg-slate-800"}
            ${status === "processing" ? "pointer-events-none opacity-70" : ""}
          `}
          onClick={() => fileInputRef.current?.click()}
          onDragOver={(e) => {
            e.preventDefault();
            setIsDragging(true);
          }}
          onDragLeave={() => setIsDragging(false)}
          onDrop={handleDrop}
        >
          <input
            ref={fileInputRef}
            type="file"
            accept=".csv,.xlsx,.xls"
            className="hidden"
            onChange={handleFileChange}
          />

          <div className="flex flex-col items-center justify-center py-16 px-6 text-center">
            {status === "processing" ? (
              <>
                <div className="w-12 h-12 border-4 border-blue-500 border-t-transparent rounded-full animate-spin mb-4" />
                <p className="text-white font-medium">
                  Converting your file...
                </p>
                <p className="text-slate-400 text-sm mt-1">
                  This only takes a moment
                </p>
              </>
            ) : status === "done" ? (
              <>
                <div className="w-14 h-14 rounded-full bg-green-500/20 border border-green-500/40 flex items-center justify-center mb-4">
                  <svg
                    className="w-7 h-7 text-green-400"
                    fill="none"
                    stroke="currentColor"
                    viewBox="0 0 24 24"
                  >
                    <path
                      strokeLinecap="round"
                      strokeLinejoin="round"
                      strokeWidth={2}
                      d="M5 13l4 4L19 7"
                    />
                  </svg>
                </div>
                <p className="text-white font-semibold text-lg">
                  Download started!
                </p>
                <p className="text-slate-400 text-sm mt-4">
                  Drop another file to convert again
                </p>
              </>
            ) : status === "error" ? (
              <>
                <div className="w-14 h-14 rounded-full bg-red-500/20 border border-red-500/40 flex items-center justify-center mb-4">
                  <svg
                    className="w-7 h-7 text-red-400"
                    fill="none"
                    stroke="currentColor"
                    viewBox="0 0 24 24"
                  >
                    <path
                      strokeLinecap="round"
                      strokeLinejoin="round"
                      strokeWidth={2}
                      d="M6 18L18 6M6 6l12 12"
                    />
                  </svg>
                </div>
                <p className="text-white font-semibold">
                  Oops, something went wrong
                </p>
                <p className="text-red-400 text-sm mt-1">{errorMsg}</p>
                <p className="text-slate-400 text-sm mt-3">
                  Click to try again
                </p>
              </>
            ) : (
              <>
                <svg
                  className="w-12 h-12 text-slate-500 mb-4"
                  fill="none"
                  stroke="currentColor"
                  viewBox="0 0 24 24"
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    strokeWidth={1.5}
                    d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"
                  />
                </svg>
                <p className="text-white font-medium text-lg">
                  Drop your CSV, XLSX, or XLS here
                </p>
                <p className="text-slate-400 text-sm mt-1">
                  or{" "}
                  <span className="text-blue-400 underline">
                    click to browse
                  </span>
                </p>
              </>
            )}
          </div>
        </div>
      </div>
    </main>
  );
}

