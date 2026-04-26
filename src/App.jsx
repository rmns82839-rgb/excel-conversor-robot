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

// Columnas que deben guardarse como texto para evitar notación científica
const TEXT_COLUMNS = [
  "orden", "identificacion", "documento", "cedula", "ticket",
  "ordenlaboratorio", "ordenfinal", "atencion", "estado"
];

function isTextColumn(colName) {
  return TEXT_COLUMNS.some(t => normalize(colName).includes(t));
}

// ─── UTILIDADES ───────────────────────────────────────────────────────────────

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

// Convierte valor a texto limpio para columnas numéricas grandes
function toCleanText(val) {
  if (val === null || val === undefined || val === "") return "";
  // Si es número, convertir a string sin notación científica
  if (typeof val === "number") {
    return val.toFixed(0);
  }
  return String(val).trim();
}

// ─── Detectar fila incompleta ─────────────────────────────────────────────────
function isRowIncomplete(newRow, targetHeaders) {
  // Para VIP: ticket=col 2 (ATENCION), orden=col 3, fecha=col 4, hora=col 6
  // Para IDIME: ticket=col 1 (TICKET), orden=col 4, fecha=col 7, hora=col 6
  // Buscar por nombre normalizado
  const ticketIdx = targetHeaders.findIndex(h => {
    const n = normalize(h);
    return n === "atencion" || n === "ticket";
  });
  const ordenIdx = targetHeaders.findIndex(h => {
    const n = normalize(h);
    return n === "ordenlaboratorio" || n === "ordenfinal";
  });
  const fechaIdx = targetHeaders.findIndex(h => {
    const n = normalize(h);
    return (n.includes("fecha") && n.includes("envio")) || n === "fechaenvioresultados";
  });
  const horaIdx = targetHeaders.findIndex(h => normalize(h) === "hora");

  const ticket = ticketIdx !== -1 ? String(newRow[ticketIdx] ?? "").trim() : "";
  const orden  = ordenIdx  !== -1 ? String(newRow[ordenIdx]  ?? "").trim() : "";
  const fecha  = fechaIdx  !== -1 ? String(newRow[fechaIdx]  ?? "").trim() : "";
  const hora   = horaIdx   !== -1 ? String(newRow[horaIdx]   ?? "").trim() : "";

  return !ticket || !orden || !fecha || !hora;
}

// ─── MAPEO DE FILAS ───────────────────────────────────────────────────────────
function mapRows(sheetData, targetHeaders, modelMap) {
  const headerIdx = findHeaderRow(sheetData);
  const sourceHeaders = sheetData[headerIdx].map((h) => String(h ?? ""));
  const mapping = buildMapping(sourceHeaders, modelMap);

  const dateColIdx = findSourceCol(sourceHeaders,
    ["fechadeenvio","fechaenvio","fechadeagendamiento","fechaingreso","fechayhora"]);

  const completeRows   = [];
  const incompleteRows = [];

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
      // Columnas de ID/ticket/orden → texto para evitar notación científica
      if (isTextColumn(col)) return toCleanText(val);
      return val;
    });

    newRow._dateVal = dateColIdx !== -1 ? srcRow[dateColIdx] : null;

    if (isRowIncomplete(newRow, targetHeaders)) {
      incompleteRows.push(newRow);
    } else {
      completeRows.push(newRow);
    }
  }

  return { completeRows, incompleteRows };
}

// ─── Crear worksheet con formato texto en columnas numéricas grandes ──────────
function makeWorksheet(headers, rows) {
  const data = [headers, ...rows];
  const ws = XLSX.utils.aoa_to_sheet(data);

  // Aplicar formato texto (@) a columnas que lo necesitan
  headers.forEach((col, colIdx) => {
    if (isTextColumn(col)) {
      const colLetter = XLSX.utils.encode_col(colIdx);
      // Aplicar a todas las filas de datos
      for (let r = 1; r < data.length; r++) {
        const cellRef = `${colLetter}${r + 1}`;
        if (ws[cellRef]) {
          ws[cellRef].t = "s"; // tipo string
          ws[cellRef].v = String(ws[cellRef].v ?? "").replace(/\.0+$/, "");
          ws[cellRef].w = ws[cellRef].v;
          delete ws[cellRef].z;
        }
      }
    }
  });

  return ws;
}

// ─── CONVERSIÓN PRINCIPAL ─────────────────────────────────────────────────────
function convertWorkbook(wb, targetModel) {
  const targetHeaders = targetModel === "IDIME" ? IDIME_HEADERS : VIP_HEADERS;
  const modelMap      = targetModel === "IDIME" ? IDIME_MAP      : VIP_MAP;
  const newWb = XLSX.utils.book_new();

  let allComplete   = [];
  let allIncomplete = [];

  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    const sheetData = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: false });
    const { completeRows, incompleteRows } = mapRows(sheetData, targetHeaders, modelMap);
    allComplete   = allComplete.concat(completeRows);
    allIncomplete = allIncomplete.concat(incompleteRows);
  }

  const cleanRow = (r) => { const c = [...r]; delete c._dateVal; return c; };

  if (targetModel === "IDIME") {
    const ws = makeWorksheet(targetHeaders, allComplete.map(cleanRow));
    XLSX.utils.book_append_sheet(newWb, ws, "IDIME");

  } else {
    const byDay = new Map();

    for (const row of allComplete) {
      const dateObj = toDateObj(row._dateVal);
      if (dateObj) {
        const key = `${dateObj.getFullYear()}-${String(dateObj.getMonth()+1).padStart(2,"0")}-${String(dateObj.getDate()).padStart(2,"0")}`;
        if (!byDay.has(key)) byDay.set(key, { date: dateObj, rows: [] });
        byDay.get(key).rows.push(row);
      } else {
        if (!byDay.has("sin-fecha")) byDay.set("sin-fecha", { date: null, rows: [] });
        byDay.get("sin-fecha").rows.push(row);
      }
    }

    const sortedDays = [...byDay.entries()]
      .filter(([k]) => k !== "sin-fecha")
      .sort((a, b) => a[0].localeCompare(b[0]));

    for (const [, { date, rows }] of sortedDays) {
      let name = dateToSheetName(date);
      let counter = 2;
      while (newWb.SheetNames.includes(name)) name = `${dateToSheetName(date)}_${counter++}`;
      const ws = makeWorksheet(targetHeaders, rows.map(cleanRow));
      XLSX.utils.book_append_sheet(newWb, ws, name);
    }

    if (byDay.has("sin-fecha")) {
      const ws = makeWorksheet(targetHeaders, byDay.get("sin-fecha").rows.map(cleanRow));
      XLSX.utils.book_append_sheet(newWb, ws, "Sin fecha");
    }

    if (newWb.SheetNames.length === 0) {
      XLSX.utils.book_append_sheet(newWb, makeWorksheet(targetHeaders, []), "Hoja1");
    }
  }

  // Hoja REVISAR al final para ambos modelos
  if (allIncomplete.length > 0) {
    const ws = makeWorksheet(targetHeaders, allIncomplete.map(cleanRow));
    XLSX.utils.book_append_sheet(newWb, ws, "⚠ REVISAR");
  }

  return { newWb, incompleteCount: allIncomplete.length };
}

// ─── COMPONENTE ───────────────────────────────────────────────────────────────
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
        const wb = XLSX.read(new Uint8Array(e.target.result), { type:"array", cellDates:true, raw:false });
        const { newWb, incompleteCount } = convertWorkbook(wb, model);
        const outName = `${file.name.replace(/\.[^.]+$/,"")}_${model}.xlsx`;
        XLSX.writeFile(newWb, outName);

        const n = newWb.SheetNames.filter(s => s !== "⚠ REVISAR").length;
        const info = model === "VIP"
          ? `${n} hoja${n!==1?"s":""} por día`
          : "1 hoja consolidada";
        const revisar = incompleteCount > 0
          ? ` · ⚠ ${incompleteCount} fila${incompleteCount!==1?"s":""} en hoja REVISAR`
          : " · ✅ Sin filas incompletas";

        setStatus("done");
        setMessage(`✅ Descargado como "${outName}" — ${info}${revisar}`);
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
              ? "📌 Filas completas → una hoja por día. Filas con ticket, fecha u hora vacíos → hoja ⚠ REVISAR."
              : "📌 Filas completas → una sola hoja IDIME. Filas con ticket, fecha u hora vacíos → hoja ⚠ REVISAR."}
          </p>
        </div>
      </main>

      <footer className="app-footer">
        Conversor Robot · Datos procesados localmente en tu navegador · Sin servidor
      </footer>
    </div>
  );
}