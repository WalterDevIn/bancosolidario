const express = require("express");
const fs = require("fs/promises");
const path = require("path");
const ExcelJS = require("exceljs");
const { nanoid } = require("nanoid");

const app = express();

// CORS libre (permite cualquier origen)
app.use((req, res, next) => {
  res.header("Access-Control-Allow-Origin", "*");
  res.header("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, OPTIONS");
  res.header("Access-Control-Allow-Headers", "Content-Type, Authorization");

  // Manejo de preflight
  if (req.method === "OPTIONS") {
    return res.sendStatus(200);
  }

  next();
});

app.use(express.json());

const DATA_DIR = path.join(__dirname, "data");
const DB_PATH = path.join(DATA_DIR, "plans.json");

// ---------- Helpers FS (sin DB) ----------
async function ensureDb() {
  await fs.mkdir(DATA_DIR, { recursive: true });
  try {
    await fs.access(DB_PATH);
  } catch {
    await fs.writeFile(DB_PATH, JSON.stringify({ plans: {} }, null, 2), "utf8");
  }
}

async function readDb() {
  await ensureDb();
  const raw = await fs.readFile(DB_PATH, "utf8");
  return JSON.parse(raw);
}

async function writeDb(db) {
  await ensureDb();
  const tmp = DB_PATH + ".tmp";
  await fs.writeFile(tmp, JSON.stringify(db, null, 2), "utf8");
  await fs.rename(tmp, DB_PATH);
}

function toMoneyInt(n) {
  // manejamos dinero como entero (sin decimales)
  const x = Number(n);
  if (!Number.isFinite(x)) throw new Error("Número inválido");
  return Math.round(x);
}

function parseISODate(s) {
  // Espera "YYYY-MM-DD"
  const d = new Date(s);
  if (Number.isNaN(d.getTime())) throw new Error("Fecha inválida (use YYYY-MM-DD)");
  // normalizamos a medianoche local
  d.setHours(0, 0, 0, 0);
  return d;
}

function fmtDateDDMMYY(d) {
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const yy = String(d.getFullYear()).slice(-2);
  return `${dd}-${mm}-${yy}`;
}

function addMonthsKeepDay(date, months) {
  // suma meses intentando mantener el día; si se pasa, cae al último día del mes
  const d = new Date(date);
  const day = d.getDate();
  d.setMonth(d.getMonth() + months, 1);
  const lastDay = new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate();
  d.setDate(Math.min(day, lastDay));
  return d;
}

// ---------- Núcleo: Generación de plan tipo foto (SAC) ----------
function generateSchedule({ monto, cuotas, tasaMensual, primeraCuotaFechaISO }) {
  const P = toMoneyInt(monto);
  const n = Number(cuotas);
  const r = Number(tasaMensual);

  if (!Number.isInteger(n) || n <= 0) throw new Error("cuotas inválido");
  if (!Number.isFinite(r) || r < 0) throw new Error("tasaMensual inválida");

  const firstDate = parseISODate(primeraCuotaFechaISO);

  // capital fijo (ajustamos última cuota por redondeo)
  const capitalBase = Math.floor(P / n);
  const capitalResto = P - capitalBase * n;

  const rows = [];
  let saldo = P;

  for (let i = 1; i <= n; i++) {
    const fecha = addMonthsKeepDay(firstDate, i - 1);
    const capital = capitalBase + (i === n ? capitalResto : 0);

    // interés sobre saldo ANTES de pagar (como en la foto)
    const interes = toMoneyInt(saldo * r);

    const total = capital + interes;

    rows.push({
      cuota: i,
      fechaISO: fecha.toISOString().slice(0, 10),
      fecha: fmtDateDDMMYY(fecha),
      saldo,
      capital,
      interes,
      total,
    });

    saldo = saldo - capital;
  }

  const sumCapital = rows.reduce((a, x) => a + x.capital, 0);
  const sumInteres = rows.reduce((a, x) => a + x.interes, 0);
  const sumTotal = rows.reduce((a, x) => a + x.total, 0);

  return { rows, sumCapital, sumInteres, sumTotal };
}

// ---------- Excel (parecido a la imagen) ----------
async function buildExcel(plan) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Plan de Pago");

  ws.properties.defaultRowHeight = 18;

  // --- Configuración de página: A4 vertical + márgenes estándar + 1 hoja ---
  ws.pageSetup = {
    paperSize: 9, // A4
    orientation: "portrait",
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 1,
    margins: {
      left: 0.7,
      right: 0.7,
      top: 0.75,
      bottom: 0.75,
      header: 0.3,
      footer: 0.3,
    },
  };

  // Columnas: N°, Fecha, Saldo, Capital, Interes, Total, Firma Solicitante, Firma Responsable
  // (NO tocar anchos como pediste)
  ws.columns = [
    { key: "cuota", width: 8 },
    { key: "fecha", width: 12 },
    { key: "saldo", width: 14 },
    { key: "capital", width: 14 },
    { key: "interes", width: 14 },
    { key: "total", width: 14 },
    { key: "firma1", width: 20 },
    { key: "firma2", width: 20 },
  ];

  const borderAll = {
    top: { style: "thin" },
    left: { style: "thin" },
    bottom: { style: "thin" },
    right: { style: "thin" },
  };

  // --- Colores de tabla ---
  const FILL_STRIPE = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF2F2F2" } }; // gris suave
  const FILL_TOTAL = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9D9D9" } }; // gris más oscuro

  function setBorderRow(rowNumber, fromCol, toCol) {
    for (let c = fromCol; c <= toCol; c++) {
      ws.getRow(rowNumber).getCell(c).border = borderAll;
    }
  }

  function fillRow(rowNumber, fromCol, toCol, fill) {
    for (let c = fromCol; c <= toCol; c++) {
      ws.getRow(rowNumber).getCell(c).fill = fill;
    }
  }

  function moneyFmt(cell) {
    cell.numFmt = '"$" #,##0';
  }

  // RichText helper: "LABEL:" en negrita + " valor" normal
  function setLabelValueCell(cell, labelWithColon, value, { boldValue = false, boldLabel = true } = {}) {
    cell.value = {
      richText: [
        { text: labelWithColon, font: { bold: !!boldLabel } },
        { text: ` ${value ?? ""}`, font: { bold: !!boldValue } },
      ],
    };
  }

  // Header superior (simulando tu hoja)
  ws.mergeCells("A1:F1");
  ws.getCell("A1").value = "GRUPO SOLIDARIO";
  ws.getCell("A1").font = { bold: true, size: 14 };
  ws.getCell("A1").alignment = { horizontal: "center" };

  ws.mergeCells("A2:F2");
  ws.getCell("A2").value = "HOY POR MI MAÑANA POR TI";
  ws.getCell("A2").font = { bold: true, size: 10 };
  ws.getCell("A2").alignment = { horizontal: "center" };

  ws.mergeCells("G1:H1");
  // (NO negrita) PLAN DE PAGO
  ws.getCell("G1").value = `PLAN DE PAGO: N° ${plan.planNumero ?? "-"}`;
  ws.getCell("G1").font = { bold: false };
  ws.getCell("G1").alignment = { horizontal: "center", vertical: "middle" };

  ws.mergeCells("G2:H2");
  // (NO negrita) "Gestión:" (pero el nombre sigue en negrita, como pediste antes)
  setLabelValueCell(ws.getCell("G2"), "Gestión:", plan.gestion ?? "-", { boldValue: true, boldLabel: false });
  ws.getCell("G2").alignment = { horizontal: "center", vertical: "middle" };

  ws.mergeCells("G4:H4");
  // "FIRMA:" en negrita (esto sí lo querías)
  ws.getCell("G4").value = { richText: [{ text: "FIRMA:", font: { bold: true } }, { text: " ____________________" }] };
  ws.getCell("G4").alignment = { horizontal: "right", vertical: "middle" };

  // Datos del solicitante (labels en negrita)
  setLabelValueCell(ws.getCell("A4"), "Nombre:", plan.nombre ?? "");
  setLabelValueCell(ws.getCell("E4"), "DNI:", plan.dni ?? "");

  setLabelValueCell(ws.getCell("A5"), "Fecha de desembolso:", plan.fechaDesembolso ?? "");
  setLabelValueCell(ws.getCell("A6"), "Monto del préstamo:", plan.monto ?? "");
  setLabelValueCell(ws.getCell("A7"), "Interés Mensual:", `${Math.round(plan.tasaMensual * 10000) / 100}%`);
  setLabelValueCell(
    ws.getCell("A8"),
    "Gastos Administrativos:",
    `${Math.round((plan.gastoAdmin ?? 0) * 10000) / 100}%`
  );
  setLabelValueCell(ws.getCell("A9"), "Tiempo de préstamo:", `${plan.cuotas} meses`);
  setLabelValueCell(ws.getCell("A10"), "Forma de Pago:", plan.formaPago ?? "Efectivo");

  // Tabla encabezados (SIN color, solo que no se corte)
  const headerRowIdx = 12;

  ws.getRow(headerRowIdx).values = [
    "N° de cuota",
    "Fecha",
    "Saldo",
    "Capital",
    "Interes",
    "Total",
    "Firma Solicitante",
    "Firma Responsable",
  ];

  const headerRow = ws.getRow(headerRowIdx);
  headerRow.font = { bold: true };
  headerRow.height = 34; // evita corte vertical
  headerRow.alignment = { horizontal: "center", vertical: "middle", wrapText: true };

  for (let c = 1; c <= 8; c++) {
    headerRow.getCell(c).alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  }

  setBorderRow(headerRowIdx, 1, 8);

  // Filas
  let r = headerRowIdx + 1;
  for (const row of plan.schedule.rows) {
    ws.getRow(r).getCell(1).value = row.cuota;
    ws.getRow(r).getCell(2).value = row.fecha;

    ws.getRow(r).getCell(3).value = row.saldo;
    moneyFmt(ws.getRow(r).getCell(3));

    ws.getRow(r).getCell(4).value = row.capital;
    moneyFmt(ws.getRow(r).getCell(4));

    ws.getRow(r).getCell(5).value = row.interes;
    moneyFmt(ws.getRow(r).getCell(5));

    ws.getRow(r).getCell(6).value = row.total;
    moneyFmt(ws.getRow(r).getCell(6));

    // Col 1 centrada H+V
    ws.getRow(r).getCell(1).alignment = { horizontal: "center", vertical: "middle" };

    // filas alternadas
    const isStripe = row.cuota % 2 === 1;
    if (isStripe) fillRow(r, 1, 8, FILL_STRIPE);

    setBorderRow(r, 1, 8);
    r++;
  }

  // --- Ajuste de alturas para que TODO entre en 1 hoja y la tabla ocupe el resto ---
  const dataStartRow = headerRowIdx + 1;
  const dataEndRow = r - 1;
  const dataRowCount = dataEndRow - dataStartRow + 1;

  const A4_HEIGHT_PTS = 841.89;
  const m = ws.pageSetup.margins;
  const marginsPts = (m.top + m.bottom + m.header + m.footer) * 72;

  const fixedRows = [];
  for (let i = 1; i <= 10; i++) fixedRows.push(i);
  fixedRows.push(11);
  fixedRows.push(headerRowIdx);
  fixedRows.push(r);

  const defaultH = ws.properties.defaultRowHeight || 18;

  function getRowHeightPts(rowNum) {
    const h = ws.getRow(rowNum).height;
    return Number.isFinite(h) ? h : defaultH;
  }

  const fixedHeightPts = fixedRows.reduce((acc, rowNum) => acc + getRowHeightPts(rowNum), 0);
  const availableForTablePts = A4_HEIGHT_PTS - marginsPts - fixedHeightPts;

  if (dataRowCount > 0) {
    let h = availableForTablePts / dataRowCount;
    h = Math.max(14, Math.min(34, h));
    for (let rr = dataStartRow; rr <= dataEndRow; rr++) {
      ws.getRow(rr).height = h;
    }
  }

  // Resumen
  ws.mergeCells(`A${r}:C${r}`);
  ws.getCell(`A${r}`).value = "RESUMEN TOTAL:";
  ws.getCell(`A${r}`).font = { bold: true };
  ws.getCell(`A${r}`).alignment = { horizontal: "left", vertical: "middle" };

  ws.getCell(`D${r}`).value = plan.schedule.sumCapital;
  moneyFmt(ws.getCell(`D${r}`));
  ws.getCell(`D${r}`).font = { bold: true };

  ws.getCell(`E${r}`).value = plan.schedule.sumInteres;
  moneyFmt(ws.getCell(`E${r}`));
  ws.getCell(`E${r}`).font = { bold: true };

  ws.getCell(`F${r}`).value = plan.schedule.sumTotal;
  moneyFmt(ws.getCell(`F${r}`));
  ws.getCell(`F${r}`).font = { bold: true };

  setBorderRow(r, 1, 8);
  fillRow(r, 1, 8, FILL_TOTAL);

  // Alineaciones finales de tabla:
  for (let rr = headerRowIdx + 1; rr < r; rr++) {
    ws.getRow(rr).getCell(1).alignment = { horizontal: "center", vertical: "middle" };

    ws.getRow(rr).getCell(2).alignment = { vertical: "middle" };
    ws.getRow(rr).getCell(7).alignment = { vertical: "middle" };
    ws.getRow(rr).getCell(8).alignment = { vertical: "middle" };

    ws.getRow(rr).getCell(3).alignment = { horizontal: "right", vertical: "middle" };
    ws.getRow(rr).getCell(4).alignment = { horizontal: "right", vertical: "middle" };
    ws.getRow(rr).getCell(5).alignment = { horizontal: "right", vertical: "middle" };
    ws.getRow(rr).getCell(6).alignment = { horizontal: "right", vertical: "middle" };
  }

  ws.getRow(r).getCell(4).alignment = { horizontal: "right", vertical: "middle" };
  ws.getRow(r).getCell(5).alignment = { horizontal: "right", vertical: "middle" };
  ws.getRow(r).getCell(6).alignment = { horizontal: "right", vertical: "middle" };

  return wb;
}

// ---------- Endpoints CRUD + Excel ----------

// Crear
app.post("/plans", async (req, res) => {
  try {
    const input = req.body;

    // defaults como la foto
    const plan = {
      id: nanoid(10),
      planNumero: input.planNumero ?? null,
      gestion: input.gestion ?? null,

      nombre: input.nombre,
      dni: input.dni,

      fechaDesembolso: input.fechaDesembolso, // "YYYY-MM-DD"
      monto: toMoneyInt(input.monto),

      tasaMensual: input.tasaMensual ?? 0.08,
      gastoAdmin: input.gastoAdmin ?? 0.005, // se guarda aunque no se use en tabla
      cuotas: input.cuotas ?? 24,
      formaPago: input.formaPago ?? "Efectivo",

      primeraCuotaFecha: input.primeraCuotaFecha, // "YYYY-MM-DD" (recomendado)
    };

    if (!plan.nombre || !plan.dni || !plan.fechaDesembolso || !plan.primeraCuotaFecha) {
      return res.status(400).json({
        error:
          "Faltan datos. Requeridos: nombre, dni, monto, fechaDesembolso (YYYY-MM-DD), primeraCuotaFecha (YYYY-MM-DD).",
      });
    }

    plan.schedule = generateSchedule({
      monto: plan.monto,
      cuotas: plan.cuotas,
      tasaMensual: plan.tasaMensual,
      primeraCuotaFechaISO: plan.primeraCuotaFecha,
    });

    const db = await readDb();
    db.plans[plan.id] = plan;
    await writeDb(db);

    res.status(201).json(plan);
  } catch (e) {
    res.status(400).json({ error: e.message ?? "Error" });
  }
});

// Leer uno
app.get("/plans/:id", async (req, res) => {
  const db = await readDb();
  const plan = db.plans[req.params.id];
  if (!plan) return res.status(404).json({ error: "No existe" });
  res.json(plan);
});

// Listar
app.get("/plans", async (req, res) => {
  const db = await readDb();
  res.json(Object.values(db.plans));
});

// Editar (recalcula schedule si cambias datos base)
app.put("/plans/:id", async (req, res) => {
  try {
    const db = await readDb();
    const existing = db.plans[req.params.id];
    if (!existing) return res.status(404).json({ error: "No existe" });

    const patch = req.body;

    const updated = {
      ...existing,
      ...patch,
    };

    // Normalizaciones
    if (patch.monto !== undefined) updated.monto = toMoneyInt(patch.monto);
    if (patch.cuotas !== undefined) updated.cuotas = Number(patch.cuotas);
    if (patch.tasaMensual !== undefined) updated.tasaMensual = Number(patch.tasaMensual);

    // Recalcular plan si toca algo clave
    const touchedKey =
      patch.monto !== undefined ||
      patch.cuotas !== undefined ||
      patch.tasaMensual !== undefined ||
      patch.primeraCuotaFecha !== undefined;

    if (touchedKey) {
      if (!updated.primeraCuotaFecha) {
        return res.status(400).json({ error: "primeraCuotaFecha requerida para recalcular" });
      }
      updated.schedule = generateSchedule({
        monto: updated.monto,
        cuotas: updated.cuotas,
        tasaMensual: updated.tasaMensual,
        primeraCuotaFechaISO: updated.primeraCuotaFecha,
      });
    }

    db.plans[req.params.id] = updated;
    await writeDb(db);

    res.json(updated);
  } catch (e) {
    res.status(400).json({ error: e.message ?? "Error" });
  }
});

// Exportar Excel
app.get("/plans/:id/excel", async (req, res) => {
  try {
    const db = await readDb();
    const plan = db.plans[req.params.id];
    if (!plan) return res.status(404).json({ error: "No existe" });

    const wb = await buildExcel(plan);

    const filename = `plan_${plan.id}.xlsx`;
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);

    await wb.xlsx.write(res);
    res.end();
  } catch (e) {
    res.status(400).json({ error: e.message ?? "Error" });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, async () => {
  await ensureDb();
  console.log(`API lista en http://localhost:${PORT}`);
});
