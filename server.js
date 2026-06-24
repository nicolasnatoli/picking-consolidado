const express = require("express");
const multer  = require("multer");
const cors    = require("cors");
const path    = require("path");
const fs      = require("fs");
const Anthropic = require("@anthropic-ai/sdk");
const ExcelJS   = require("exceljs");

const app    = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20 * 1024 * 1024 } });
const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, "public")));

// ─── CONSTANTS ────────────────────────────────────────────────────────────────
const PROMPT = `Extraé los datos de esta orden de armado industrial argentina.
Respondé ÚNICAMENTE con JSON válido. Sin texto antes ni después. Sin markdown.

Ejemplo del formato esperado:
{
  "producto": "1401",
  "descripcion_producto": "ESPEJO MB ACCELO MANUAL",
  "comprobante": "2937",
  "fecha": "17/06/2026",
  "items": [
    {"insumo": "0103033", "descripcion": "VIDRIO CONVEXO ESPEJO 1400", "cantidad": 10, "um": "UNI", "ubicacion": "PB-A-1-2"},
    {"insumo": "0202010", "descripcion": "CINTA AFT 5369 3M", "cantidad": 400, "um": "CMT", "ubicacion": "PB-C-12-5"},
    {"insumo": "0800006", "descripcion": "MO ARMADO", "cantidad": 3890, "um": "SEG", "ubicacion": ""}
  ]
}

REGLAS CRÍTICAS:
- comprobante: últimos 4 dígitos del número (005 00005-00002937 → "2937")
- cantidad: número, no string. Si dice 400.00 → 400
- ubicacion: EXACTAMENTE como figura en col Ubic. Si no tiene → ""
- Incluí ABSOLUTAMENTE TODOS los ítems sin excepción
- Si hay múltiples órdenes, devolvé array: [{orden1}, {orden2}]
- Si hay una sola orden, devolvé el objeto directamente`;

const ZONE_NAMES = {
  0: "ZONA 2P — SEGUNDO PISO",
  1: "ZONA 1P — PRIMER PISO",
  2: "ZONA PB — PLANTA BAJA",
  3: "ZONA P  — PASILLOS",
  4: "ZONA M  — MÓDULOS M",
  5: "OTRAS UBICACIONES",
  6: "SIN UBICACIÓN"
};

const ZONE_COLORS = {
  0:"1E3A5F", 1:"14532D", 2:"3B1F5E",
  3:"7C2D12", 4:"374151", 5:"1F1F1F", 6:"333333"
};

const GRP_COLORS = [
  "2563EB","D97706","059669","DC2626","7C3AED",
  "0891B2","65A30D","DB2777","EA580C","0D9488",
  "1D4ED8","B45309","047857","B91C1C","6D28D9",
  "0E7490","4D7C0F","BE185D","C2410C","0F766E"
];

// ─── HELPERS ─────────────────────────────────────────────────────────────────
function zoneKey(ubic) {
  const u = (ubic || "").trim().toUpperCase();
  if (!u)              return 6;
  if (u.startsWith("2P")) return 0;
  if (u.startsWith("1P")) return 1;
  if (u.startsWith("PB")) return 2;
  if (u.startsWith("P"))  return 3;
  if (u.startsWith("M"))  return 4;
  return 5;
}

function extractJSON(text) {
  const clean = text.replace(/```json|```/gi, "").trim();
  for (const pattern of [/(\[[\s\S]*\])/, /(\{[\s\S]*\})/]) {
    const m = clean.match(pattern);
    if (m) {
      try {
        const p = JSON.parse(m[1]);
        return Array.isArray(p) ? p : [p];
      } catch {}
    }
  }
  throw new Error("No se pudo extraer JSON de la respuesta");
}

// ─── API: PROCESAR IMÁGENES ───────────────────────────────────────────────────
app.post("/api/procesar", upload.array("archivos", 30), async (req, res) => {
  if (!req.files || req.files.length === 0)
    return res.status(400).json({ error: "No se recibieron archivos" });

  const results = [];
  let grupo = 0;

  for (const file of req.files) {
    try {
      const isPDF  = file.mimetype === "application/pdf";
      const isHEIC = file.mimetype === "image/heic" ||
                     file.mimetype === "image/heif" ||
                     file.originalname.toLowerCase().endsWith(".heic") ||
                     file.originalname.toLowerCase().endsWith(".heif");

      let imageMime = file.mimetype;
      if (isHEIC || !["image/jpeg","image/png","image/gif","image/webp"].includes(imageMime)) {
        imageMime = "image/jpeg";
      }

      const b64 = file.buffer.toString("base64");

      const contentBlock = isPDF
        ? { type: "document", source: { type: "base64", media_type: "application/pdf", data: b64 } }
        : { type: "image",    source: { type: "base64", media_type: imageMime,          data: b64 } };

      const resp = await client.messages.create({
        model:      "claude-sonnet-4-6",
        max_tokens: 4000,
        messages:   [{ role: "user", content: [contentBlock, { type: "text", text: PROMPT }] }]
      });

      const text   = resp.content.find(b => b.type === "text")?.text || "";
      const orders = extractJSON(text);

      for (const order of orders) {
        grupo++;
        results.push({ grupo, order, filename: file.originalname });
      }
    } catch (e) {
      results.push({ grupo: ++grupo, error: e.message, filename: file.originalname });
    }
  }

  res.json({ results });
});

// ─── API: GENERAR EXCEL ───────────────────────────────────────────────────────
app.post("/api/excel", express.json({ limit: "10mb" }), async (req, res) => {
  const { results } = req.body;
  if (!results || !results.length)
    return res.status(400).json({ error: "Sin datos" });

  // Construir filas
  const allRows = [];
  const ordersInfo = [];

  for (const r of results) {
    if (r.error || !r.order) continue;
    const { grupo, order } = r;
    for (const item of (order.items || [])) {
      allRows.push({
        grupo,
        insumo:      item.insumo      || "",
        descripcion: item.descripcion || "",
        cantidad:    item.cantidad    ?? 0,
        um:          item.um          || "",
        ubicacion:   item.ubicacion   || "",
        producto:    order.producto   || "",
        fecha:       order.fecha      || "",
        of:          order.comprobante|| "",
      });
    }
    ordersInfo.push({
      grupo,
      producto: order.producto || "",
      desc:     order.descripcion_producto || "",
      of:       order.comprobante || "",
      fecha:    order.fecha || "",
      items:    (order.items || []).length
    });
  }

  // Ordenar por zona
  allRows.sort((a, b) => {
    const za = zoneKey(a.ubicacion), zb = zoneKey(b.ubicacion);
    if (za !== zb) return za - zb;
    return (a.ubicacion || "").localeCompare(b.ubicacion || "");
  });

  // Crear Excel
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Consolidado Picking");

  ws.views = [{ showGridLines: false }];

  const DARK = "1A1A2E", GOLD = "C9A84C", WHITE = "FFFFFF";
  const colWidths = [6, 6, 14, 46, 10, 7, 20, 10, 12, 7];
  const colHdrs   = ["#","Gr.","Cód. Insumo","Descripción","Cantidad","U/M","Ubicación","Producto","Fecha","OF"];

  colWidths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });

  // Título
  ws.mergeCells("A1:J1");
  const t1 = ws.getCell("A1");
  t1.value     = "CONSOLIDADO DE PICKING — ÓRDENES DE ARMADO";
  t1.font      = { name: "Arial", bold: true, size: 13, color: { argb: "FF" + WHITE } };
  t1.fill      = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + DARK } };
  t1.alignment = { horizontal: "center", vertical: "middle" };
  ws.getRow(1).height = 26;

  // Subtítulo
  const subtitle = ordersInfo.map(o => `GR.${String(o.grupo).padStart(2,"0")} · Prod.${o.producto} · OF ${o.of} · ${o.fecha}`).join("   |   ");
  ws.mergeCells("A2:J2");
  const t2 = ws.getCell("A2");
  t2.value     = subtitle;
  t2.font      = { name: "Arial", size: 8, color: { argb: "FFAAAAAA" } };
  t2.fill      = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + DARK } };
  t2.alignment = { horizontal: "center", vertical: "middle" };
  ws.getRow(2).height = 14;

  // Espacio
  ws.mergeCells("A3:J3");
  ws.getCell("A3").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + DARK } };
  ws.getRow(3).height = 4;

  // Headers
  const hRow = ws.getRow(4);
  hRow.height = 20;
  colHdrs.forEach((h, i) => {
    const c = hRow.getCell(i + 1);
    c.value     = h;
    c.font      = { name: "Arial", bold: true, size: 9, color: { argb: "FF" + WHITE } };
    c.fill      = { type: "pattern", pattern: "solid", fgColor: { argb: "FF2D2D4E" } };
    c.alignment = { horizontal: h === "Descripción" ? "left" : "center", vertical: "middle" };
    c.border    = { top:{style:"thin",color:{argb:"FF3D3D6E"}}, bottom:{style:"thin",color:{argb:"FF3D3D6E"}},
                    left:{style:"thin",color:{argb:"FF3D3D6E"}}, right:{style:"thin",color:{argb:"FF3D3D6E"}} };
  });
  ws.views = [{ state: "frozen", ySplit: 4 }];

  // Datos
  let currentZone = null, rowNum = 5, dataN = 0;
  const thinBorder = (color = "CCCCCC") => ({
    top:{style:"thin",color:{argb:"FF"+color}}, bottom:{style:"thin",color:{argb:"FF"+color}},
    left:{style:"thin",color:{argb:"FF"+color}}, right:{style:"thin",color:{argb:"FF"+color}}
  });

  for (const item of allRows) {
    const zk = zoneKey(item.ubicacion);
    const gc  = GRP_COLORS[(item.grupo - 1) % GRP_COLORS.length];
    const bg  = item.grupo % 2 === 0 ? "FFFFFF" : "F7F7F7";

    if (zk !== currentZone) {
      currentZone = zk;
      ws.mergeCells(`A${rowNum}:J${rowNum}`);
      const zc = ws.getCell(`A${rowNum}`);
      zc.value     = `▶  ${ZONE_NAMES[zk]}`;
      zc.font      = { name: "Arial", bold: true, italic: true, size: 9, color: { argb: "FF" + WHITE } };
      zc.fill      = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + (ZONE_COLORS[zk] || "333333") } };
      zc.alignment = { horizontal: "left", vertical: "middle", indent: 1 };
      ws.getRow(rowNum).height = 14;
      rowNum++;
    }

    dataN++;
    const row = ws.getRow(rowNum);
    row.height  = 15;

    const setCell = (col, value, align, bold, color, bgColor, leftAccent) => {
      const c = row.getCell(col);
      c.value     = value;
      c.font      = { name: "Arial", size: 9, bold: bold || false, color: { argb: "FF" + (color || DARK) } };
      c.fill      = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + (bgColor || bg) } };
      c.alignment = { horizontal: align || "center", vertical: "middle" };
      c.border    = leftAccent
        ? { left:{style:"medium",color:{argb:"FF"+gc}}, right:thinBorder().right, top:thinBorder().top, bottom:thinBorder().bottom }
        : thinBorder();
    };

    setCell(1,  String(dataN).padStart(2,"0"), "center", false, "888888", "F7F7F7");
    setCell(2,  String(item.grupo).padStart(2,"0"), "center", true, WHITE, gc);
    setCell(3,  item.insumo,      "center", false, DARK, bg, true);
    setCell(4,  item.descripcion, "left",   false, DARK, bg);
    setCell(5,  item.cantidad,    "right",  true,  DARK, bg);
    setCell(6,  item.um,          "center", false, "555555", bg);
    setCell(7,  item.ubicacion || "—", "center", true, DARK, bg);
    setCell(8,  item.producto,    "center", true,  DARK, bg);
    setCell(9,  item.fecha,       "center", false, "555555", bg);
    setCell(10, item.of,          "center", true,  DARK, bg);

    rowNum++;
  }

  // Pie
  ws.mergeCells(`A${rowNum}:J${rowNum}`);
  const pie = ws.getCell(`A${rowNum}`);
  pie.value     = `Total: ${dataN} ítems  ·  ${ordersInfo.length} órdenes  ·  Generado: ${new Date().toLocaleDateString("es-AR")}`;
  pie.font      = { name: "Arial", bold: true, size: 9, color: { argb: "FF" + DARK } };
  pie.fill      = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + GOLD } };
  pie.alignment = { horizontal: "center", vertical: "middle" };
  ws.getRow(rowNum).height = 16;

  // Hoja resumen
  const ws2 = wb.addWorksheet("Resumen");
  ws2.views = [{ showGridLines: false }];
  [7,10,36,10,14,8,8].forEach((w,i) => ws2.getColumn(i+1).width = w);
  ws2.mergeCells("A1:G1");
  const rs1 = ws2.getCell("A1");
  rs1.value = "RESUMEN DE ÓRDENES"; rs1.font = { name:"Arial", bold:true, size:12, color:{argb:"FF"+WHITE} };
  rs1.fill  = { type:"pattern", pattern:"solid", fgColor:{argb:"FF"+DARK} };
  rs1.alignment = { horizontal:"center", vertical:"middle" };
  ws2.getRow(1).height = 22;
  ["Grupo","Producto","Descripción","OF","Fecha","Ítems"].forEach((h,i) => {
    const c = ws2.getRow(2).getCell(i+1);
    c.value = h; c.font = { name:"Arial", bold:true, size:9, color:{argb:"FF"+WHITE} };
    c.fill  = { type:"pattern", pattern:"solid", fgColor:{argb:"FF2D2D4E"} };
    c.alignment = { horizontal:"center", vertical:"middle" }; c.border = thinBorder();
  });
  ws2.getRow(2).height = 18;
  ordersInfo.forEach((o,i) => {
    const r  = ws2.getRow(i+3);
    const gc = GRP_COLORS[(o.grupo-1) % GRP_COLORS.length];
    const bg = i%2===0 ? "F9F9F9" : "FFFFFF";
    [String(o.grupo).padStart(2,"0"), o.producto, o.desc, o.of, o.fecha, o.items].forEach((v,j) => {
      const c = r.getCell(j+1);
      c.value = v;
      c.font  = { name:"Arial", size:9, bold:j===0, color:{argb: j===0 ? "FF"+WHITE : "FF"+DARK} };
      c.fill  = { type:"pattern", pattern:"solid", fgColor:{argb: j===0 ? "FF"+gc : "FF"+bg} };
      c.alignment = { horizontal:"center", vertical:"middle" }; c.border = thinBorder();
    });
    r.height = 16;
  });

  // Enviar archivo
  const fecha = new Date().toLocaleDateString("es-AR").replace(/\//g,"");
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", `attachment; filename="Consolidado_Picking_${fecha}.xlsx"`);
  await wb.xlsx.write(res);
  res.end();
});

// Servir frontend
app.get("*", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "index.html"));
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`✅ Picking app corriendo en puerto ${PORT}`));
