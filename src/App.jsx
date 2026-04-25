import { useState, useCallback } from "react";
import * as XLSX from "xlsx";
import "./App.css";

const IDIME_HEADERS = [
  "ESTADO", "TICKET", "DOCUMENTO", "CODIGO",
  "ORDEN\nLABORATORIO", "ORDEN\nFINAL", "HORA",
  "FECHA ENVIO\nRESULTADOS", "FECHA DE\nAGENDAMIENTO",
  "NOMBRE\nPACIENTE", "DIRECCION Y\nTELEFONO DEL\nPACIENTE",
  "TIPO\nCLIENTE", "CIUDAD", "OBSERVACIÓN ", "CORREO",
];

const VIP_HEADERS = [
  "Orden", "Identificacion", "ATENCION", "orden laboratorio",
  "FECHA DE ENVIO", "FECHA DE ENVIO", "HORA", "Nombre", "Edad",
  "Examen", "Correo", "Envío", "Fecha y hora",
  "Plan ", "Responsable ", "Comentario ", "Servicio",
];

const IDIME_MAP = {
  "ESTADO":                               ["estado", "atencion"],
  "TICKET":                               ["ticket", "orden", "ticke"],
  "DOCUMENTO":                            ["documento", "identificacion", "cedula"],
  "CODIGO":                               ["codigo", "examen"],
  "ORDEN\nLABORATORIO":                   ["ordenlaboratorio", "ordenlaboratorio"],
  "ORDEN\nFINAL":                         ["ordenfinal", "ordenfinal"],
  "HORA":                                 ["hora"],
  "FECHA ENVIO\nRESULTADOS":             ["fechadeenvio", "fechaenvio", "fechayhoraenvio", "fechayhora"],
  "FECHA DE\nAGENDAMIENTO":              ["fechaingreso", "fechadeagendamiento"],
  "NOMBRE\nPACIENTE":                    ["nombre", "nombrepaciente"],
  "DIRECCION Y\nTELEFONO DEL\nPACIENTE": ["direccion", "telefono"],
  "TIPO\nCLIENTE":                        ["plan", "tipocliente"],
  "CIUDAD":                               ["ciudad", "servicio"],
  "OBSERVACIÓN ":                         ["observacion", "comentario"],
  "CORREO":                               ["correo"],
};

const VIP_MAP = {
  "Orden":             ["orden", "ticket"],
  "Identificacion":    ["identificacion", "documento", "cedula"],
  "ATENCION":          ["atencion", "estado"],
  "orden laboratorio": ["ordenlaboratorio", "ordenlaboratorio"],
  "FECHA DE ENVIO":    ["fechadeenvio", "fechaenvio"],
  "HORA":              ["hora"],
  "Nombre":            ["nombre", "nombrepaciente"],
  "Edad":              ["edad"],
  "Examen":            ["examen", "codigo"],
  "Correo":            ["correo"],
  "Envío":             ["envio", "envio"],
  "Fecha y hora":      ["fechayhora", "fechayhoraenvio"],
  "Plan ":             ["plan", "tipocliente"],
  "Responsable ":      ["responsable"],
  "Comentario ":       ["comentario", "observacion"],
  "Servicio":          ["servicio", "ciudad"],
};

function normalize(str) {
  return String(str ?? "")
    .toLowerCase().replace(/\s+/g, "")
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function findSourceCol(sourceHeaders, candidates) {
  for (const c of candidates) {
    const idx = sourceHeaders.findIndex((h) => normalize(h) === normalize(c));
    if (idx !== -1) return idx;
  }
  for (const c of candidates) {
    const idx = sourceHeaders.findIndex(
      (h) => normalize(h).includes(normalize(c)) || normalize(c).includes(normalize(h))
    );
    if (idx !== -1) return idx;
  }
  return -1;
}

function buildMapping(sourceHeaders, modelMap) {
  const result = {};
  for (const [target, candidates] of Object.entries(modelMap)) {
    result[target] = findSourceCol(sourceHeaders, candidates);
  }
  return result;
}

function findHeaderRow(sheetData) {
  const keywords = ["estado","ticket","orden","atencion","identificacion","documento","nombre","hora","fecha"];
  for (let r = 0; r < Math.min(sheetData.length, 15); r++) {
    const row = sheetData[r].map((v) => normalize(String(v ?? "")));
    const hits = keywords.filter((k) => row.some((cell) => cell.includes(k)));
    if (hits.length >= 3) return r;
  }
  return 0;
}

function toTimeString(val) {
  if (val === null || val === undefined || val === "") return "";
  if (typeof val === "string") {
    const m = val.match(/^(\d{1,2}):(\d{2})/);
    if (m) return `${m[1].padStart(2,"0")}:${m[2]}`;
    return val;
  }
  if (val instanceof Date) {
    return `${String(val.getHours()).padStart(2,"0")}:${String(val.getMinutes()).padStart(2,"0")}`;
  }
  if (typeof val === "number") {
    const frac = val % 1;
    const mins = Math.round(frac * 24 * 60);
    return `${String(Math.floor(mins/60)).padStart(2,"0")}:${String(mins%60).padStart(2,"0")}`;
  }
  return String(val);
}

function toDateString(val) {
  if (val === null || val === undefined || val === "") return "";
  if (val instanceof Date) {
    const d = String(val.getDate()).padStart(2,"0");
    const m = String(val.getMonth()+1).padStart(2,"0");
    return `${d}/${m}/${val.getFullYear()}`;
  }
  return String(val);
}

function toDateObj(val) {
  if (!val) return null;
  if (val instanceof Date && !isNaN(val)) return val;
  try { const d = new Date(val); return isNaN(d) ? null : d; } catch { return null; }
}

function dateToSheetName(date) {
  if (!date) return "Sin fecha";
  const months = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"];
  return `${String(date.getDate()).padStart(2,"0")} ${months[date.getMonth()]}`;
}

function mapRows(sheetData, targetHeaders, modelMap) {
  const headerIdx = findHeaderRow(sheetData);
  const sourceHeaders = sheetData[headerIdx].map((h) => String(h ?? ""));
  const mapping = buildMapping(sourceHeaders, modelMap);

  // Find best date column for grouping
  const dateColIdx = findSourceCol(sourceHeaders,
    ["fechadeenvio","fechaenvio","fechadeagendamiento","fechaingreso","fechayhora"]);

  const dataRows = [];
  for (let r = headerIdx + 1; r < sheetData.length; r++) {
    const srcRow = sheetData[r];
    if (!srcRow || srcRow.every((v) => v === "" || v === null || v === undefined)) continue;
    const newRow = targetHeaders.map((col) => {
      const srcIdx = mapping[col];
      if (srcIdx === -1 || srcIdx === undefined) return "";
      const val = srcRow[srcIdx] ?? "";
      const n = normalize(col);
      if (n === "hora") return toTimeString(val);
      if (n.includes("fecha")) return toDateString(val);
      return val;
    });
    newRow._dateVal = dateColIdx !== -1 ? srcRow[dateColIdx] : null;
    dataRows.push(newRow);
  }
  return dataRows;
}

function convertWorkbook(wb, targetModel) {
  const targetHeaders = targetModel === "IDIME" ? IDIME_HEADERS : VIP_HEADERS;
  const modelMap      = targetModel === "IDIME" ? IDIME_MAP      : VIP_MAP;
  const newWb = XLSX.utils.book_new();

  let allRows = [];
  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    const sheetData = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: true });
    allRows = allRows.concat(mapRows(sheetData, targetHeaders, modelMap));
  }

  const cleanRow = (r) => { const c = [...r]; delete c._dateVal; return c; };

  if (targetModel === "IDIME") {
    // All rows → single sheet
    const ws = XLSX.utils.aoa_to_sheet([targetHeaders, ...allRows.map(cleanRow)]);
    XLSX.utils.book_append_sheet(newWb, ws, "IDIME");

  } else {
    // Split by date → one sheet per day
    const byDay = new Map();
    const noDate = [];

    for (const row of allRows) {
      const dateObj = toDateObj(row._dateVal);
      if (dateObj) {
        const key = `${dateObj.getFullYear()}-${String(dateObj.getMonth()+1).padStart(2,"0")}-${String(dateObj.getDate()).padStart(2,"0")}`;
        if (!byDay.has(key)) byDay.set(key, { date: dateObj, rows: [] });
        byDay.get(key).rows.push(row);
      } else {
        noDate.push(row);
      }
    }

    const sortedDays = [...byDay.entries()].sort((a, b) => a[0].localeCompare(b[0]));

    for (const [, { date, rows }] of sortedDays) {
      let name = dateToSheetName(date);
      let counter = 2;
      while (newWb.SheetNames.includes(name)) name = `${dateToSheetName(date)}_${counter++}`;
      const ws = XLSX.utils.aoa_to_sheet([targetHeaders, ...rows.map(cleanRow)]);
      XLSX.utils.book_append_sheet(newWb, ws, name);
    }

    if (noDate.length > 0) {
      const ws = XLSX.utils.aoa_to_sheet([targetHeaders, ...noDate.map(cleanRow)]);
      XLSX.utils.book_append_sheet(newWb, ws, "Sin fecha");
    }

    if (newWb.SheetNames.length === 0) {
      XLSX.utils.book_append_sheet(newWb, XLSX.utils.aoa_to_sheet([targetHeaders]), "Hoja1");
    }
  }

  return newWb;
}

export default function App() {
  const [file, setFile]         = useState(null);
  const [model, setModel]       = useState("VIP");
  const [status, setStatus]     = useState(null);
  const [message, setMessage]   = useState("");
  const [dragging, setDragging] = useState(false);

  const handleFile = (f) => {
    if (!f) return;
    const ext = f.name.split(".").pop().toLowerCase();
    if (!["xlsx","xls"].includes(ext)) {
      setStatus("error"); setMessage("Solo se aceptan archivos .xlsx o .xls"); return;
    }
    setFile(f); setStatus(null); setMessage("");
  };

  const handleDrop = useCallback((e) => {
    e.preventDefault(); setDragging(false); handleFile(e.dataTransfer.files[0]);
  }, []);

  const handleConvert = () => {
    if (!file) { setStatus("error"); setMessage("Por favor sube un archivo primero."); return; }
    setStatus("loading"); setMessage("Convirtiendo...");
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target.result), { type:"array", cellDates:true, raw:true });
        const newWb = convertWorkbook(wb, model);
        const outName = `${file.name.replace(/\.[^.]+$/,"")}_${model}.xlsx`;
        XLSX.writeFile(newWb, outName);
        const n = newWb.SheetNames.length;
        const info = model === "VIP" ? `${n} hoja${n!==1?"s":""} (una por día)` : "1 hoja consolidada";
        setStatus("done"); setMessage(`✅ Descargado como "${outName}" — ${info}`);
      } catch (err) {
        setStatus("error"); setMessage("Error: " + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  return (
    <div className="app">
      <header className="app-header">
        <div className="logo">⚙️</div>
        <h1>Conversor de Excel</h1>
        <p className="subtitle">Transforma cualquier Excel al modelo <strong>IDIME</strong> o <strong>VIP</strong></p>
      </header>

      <main className="app-main">
        <div className="card">
          <h2>1. Selecciona el modelo destino</h2>
          <div className="model-selector">
            <button className={`model-btn ${model==="VIP"?"active":""}`} onClick={()=>setModel("VIP")}>
              🟦 Modelo VIP
              <span className="model-hint">Una hoja por día</span>
            </button>
            <button className={`model-btn ${model==="IDIME"?"active idime":""}`} onClick={()=>setModel("IDIME")}>
              🟩 Modelo IDIME
              <span className="model-hint">Todo en una hoja</span>
            </button>
          </div>
        </div>

        <div className="card">
          <h2>2. Sube tu archivo Excel</h2>
          <div
            className={`dropzone ${dragging?"dragging":""} ${file?"has-file":""}`}
            onDragOver={(e)=>{e.preventDefault();setDragging(true);}}
            onDragLeave={()=>setDragging(false)}
            onDrop={handleDrop}
            onClick={()=>document.getElementById("file-input").click()}
          >
            {file ? (
              <><span className="file-icon">📊</span>
                <p className="file-name">{file.name}</p>
                <p className="file-hint">Clic para cambiar archivo</p></>
            ) : (
              <><span className="file-icon">📂</span>
                <p>Arrastra tu archivo aquí<br/>o <strong>haz clic para buscar</strong></p>
                <p className="file-hint">Formatos: .xlsx, .xls — cualquier modelo</p></>
            )}
          </div>
          <input id="file-input" type="file" accept=".xlsx,.xls"
            style={{display:"none"}} onChange={(e)=>handleFile(e.target.files[0])}/>
        </div>

        <div className="card">
          <h2>3. Convierte y descarga</h2>
          <button className="convert-btn" onClick={handleConvert} disabled={!file||status==="loading"}>
            {status==="loading" ? "⏳ Procesando..." : `🔄 Convertir a ${model}`}
          </button>
          {status && <div className={`status-msg ${status}`}>{message}</div>}
        </div>

        <div className="card info-card">
          <h2>Columnas del modelo {model}</h2>
          <div className="col-list">
            {(model==="IDIME"?IDIME_HEADERS:VIP_HEADERS).map((h,i)=>(
              <span key={i} className="col-tag">{h.replace(/\n/g," ")}</span>
            ))}
          </div>
          <p className="info-note">
            {model==="VIP"
              ? "📌 El resultado tendrá una hoja por cada día distinto según la columna de fecha."
              : "📌 Sin importar cuántas hojas tenga el archivo, todo queda consolidado en una sola hoja."}
          </p>
        </div>
      </main>

      <footer className="app-footer">
        Conversor Robot · Datos procesados localmente en tu navegador · Sin servidor
      </footer>
    </div>
  );
}