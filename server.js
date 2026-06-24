const express  = require("express");
const multer   = require("multer");
const cors     = require("cors");
const path     = require("path");
const sharp    = require("sharp");
const Anthropic = require("@anthropic-ai/sdk");
const ExcelJS   = require("exceljs");

const app    = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20 * 1024 * 1024 } });
const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, "public")));

const PROMPT = `Extrae los datos de esta orden de armado industrial argentina.
Responde UNICAMENTE con JSON valido. Sin texto antes ni despues. Sin markdown.

Formato obligatorio:
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

REGLAS:
- comprobante: ultimos 4 digitos (005 00005-00002937 = "2937")
- cantidad: numero, no string
- ubicacion: exactamente como figura en col Ubic, o "" si no tiene
- Incluir TODOS los items sin excepcion
- Si hay multiples ordenes: array [{orden1},{orden2}]
- Si hay una sola orden: objeto directo`;

const ZONE_NAMES = {
  0:"ZONA 2P - SEGUNDO PISO", 1:"ZONA 1P - PRIMER PISO",
  2:"ZONA PB - PLANTA BAJA",  3:"ZONA P  - PASILLOS",
  4:"ZONA M  - MODULOS M",    5:"OTRAS UBICACIONES", 6:"SIN UBICACION"
};
const ZONE_COLORS = {
  0:"1E3A5F",1:"14532D",2:"3B1F5E",3:"7C2D12",4:"374151",5:"1F1F1F",6:"333333"
};
const GRP_COLORS = [
  "2563EB","D97706","059669","DC2626","7C3AED",
  "0891B2","65A30D","DB2777","EA580C","0D9488",
  "1D4ED8","B45309","047857","B91C1C","6D28D9",
  "0E7490","4D7C0F","BE185D","C2410C","0F766E"
];

function zoneKey(ubic) {
  const u = (ubic || "").trim().toUpperCase();
  if (!u) return 6;
  if (u.startsWith("2P")) return 0;
  if (u.startsWith("1P")) return 1;
  if (u.startsWith("PB")) return 2;
  if (u.startsWith("P"))  return 3;
  if (u.startsWith("M"))  return 4;
  return 5;
}

function extractJSON(text) {
  const clean = text.replace(/```json|```/gi, "").trim();
  for (const pat of [/(\[[\s\S]*\])/, /(\{[\s\S]*\})/]) {
    const m = clean.match(pat);
    if (m) {
      try {
        const p = JSON.parse(m[1]);
        return Array.isArray(p) ? p : [p];
      } catch {}
    }
  }
  throw new Error("No se pudo extraer JSON");
}

async function toJpegBuffer(buffer, mimetype, filename) {
  const name = (filename || "").toLowerCase();
  const isHEIC = mimetype === "image/heic" || mimetype === "image/heif" ||
                 name.endsWith(".heic") || name.endsWith(".heif");
  const isImage = mimetype.startsWith("image/") || isHEIC;
  const isPDF   = mimetype === "application/pdf";

  if (isPDF) return { buffer, mime: "application/pdf" };

  if (isHEIC) {
    const jpeg = await sharp(buffer).jpeg({ quality: 90 }).toBuffer();
    return { buffer: jpeg, mime: "image/jpeg" };
  }

  if (isImage && mimetype !== "image/jpeg" && mimetype !== "image/png" &&
      mimetype !== "image/gif" && mimetype !== "image/webp") {
    const jpeg = await sharp(buffer).jpeg({ quality: 90 }).toBuffer();
    return { buffer: jpeg, mime: "image/jpeg" };
  }

  return { buffer, mime: mimetype };
}

app.post("/api/procesar", upload.array("archivos", 30), async (req, res) => {
  if (!req.files || !req.files.length)
    return res.status(400).json({ error: "No se recibieron archivos" });

  const results = [];
  let grupo = 0;

  for (const file of req.files) {
    try {
      const { buffer, mime } = await toJpegBuffer(file.buffer, file.mimetype, file.originalname);
      const b64 = buffer.toString("base64");
      const isPDF = mime === "application/pdf";

      const contentBlock = isPDF
        ? { type: "document", source: { type: "base64", media_type: "application/pdf", data: b64 } }
        : { type: "image",    source: { type: "base64", media_type: mime, data: b64 } };

      const resp = await client.messages.create({
        model: "claude-sonnet-4-6",
        max_tokens: 4000,
        messages: [{ role: "user", content: [contentBlock, { type: "text", text: PROMPT }] }]
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

app.post("/api/excel", express.json({ limit: "10mb" }), async (req, res) => {
  const { results } = req.body;
  if (!results || !results.length)
    return res.status(400).json({ error: "Sin datos" });

  const allRows = [], ordersInfo = [];

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

  allRows.sort((a, b) => {
    const za = zoneKey(a.ubicacion), zb = zoneKey(b.ubicacion);
    if (za !== zb) return za - zb;
    return (a.ubicacion || "").localeCompare(b.ubicacion || "");
  });

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Consolidado Picking");
  ws.views = [{ showGridLines: false }];

  const DARK = "1A1A2E", GOLD = "C9A84C", WHITE = "FFFFFF";
  [6,6,14,46,10,7,20,10,12,7].forEach((w,i) => { ws.getColumn(i+1).width = w; });

  ws.mergeCells("A1:J1");
  Object.assign(ws.getCell("A1"), {
    value: "CONSOLIDADO DE PICKING - ORDENES DE ARMADO",
    font: { name:"Arial", bold:true, size:13, color:{argb:"FF"+WHITE} },
    fill: { type:"pattern", pattern:"solid", fgColor:{argb:"FF"+DARK} },
    alignment: { horizontal:"center", vertical:"middle" }
  });
  ws.getRow(1).height = 26;

  ws.mergeCells("A2:J2");
  Object.assign(ws.getCell("A2"), {
    value: ordersInfo.map(o => `GR.${String(o.grupo).padStart(2,"0")} Prod.${o.producto} OF ${o.of} ${o.fecha}`).join("  |  "),
    font: { name:"Arial", size:8, color:{argb:"FFAAAAAA"} },
    fill: { type:"pattern", pattern:"solid", fgColor:{argb:"FF"+DARK} },
    alignment: { horizontal:"center", vertical:"middle" }
  });
  ws.getRow(2).height = 14;

  ws.mergeCells("A3:J3");
  ws.getCell("A3").fill = { type:"pattern", pattern:"solid", fgColor:{argb:"FF"+DARK} };
  ws.getRow(3).height = 4;

  const hdrs = ["#","Gr.","Cod. Insumo","Descripcion","Cantidad","U/M","Ubicacion","Producto","Fecha","OF"];
  hdrs.forEach((h,i) => {
    const c = ws.getRow(4).getCell(i+1);
    c.value = h;
    c.font = { name:"Arial", bold:true, size:9, color:{argb:"FF"+WHITE} };
    c.fill = { type:"pattern", pattern:"solid", fgColor:{argb:"FF2D2D4E"} };
    c.alignment = { horizontal: h==="Descripcion"?"left":"center", vertical:"middle" };
    c.border = {
      top:{style:"thin",color:{argb:"FF3D3D6E"}},
      bottom:{style:"thin",color:{argb:"FF3D3D6E"}},
      left:{style:"thin",color:{argb:"FF3D3D6E"}},
      right:{style:"thin",color:{argb:"FF3D3D6E"}}
    };
  });
  ws.getRow(4).height = 20;
  ws.views = [{ state:"frozen", ySplit:4 }];

  let currentZone = null, rowNum = 5, dataN = 0;

  const tb = () => ({
    top:{style:"thin",color:{argb:"FFCCCCCC"}},
    bottom:{style:"thin",color:{argb:"FFCCCCCC"}},
    left:{style:"thin",color:{argb:"FFCCCCCC"}},
    right:{style:"thin",color:{argb:"FFCCCCCC"}}
  });

  for (const item of allRows) {
    const zk = zoneKey(item.ubicacion);
    const gc = GRP_COLORS[(item.grupo-1) % GRP_COLORS.length];
    const bg = item.grupo % 2 === 0 ? "FFFFFF" : "F7F7F7";

    if (zk !== currentZone) {
      currentZone = zk;
      ws.mergeCells(`A${rowNum}:J${rowNum}`);
      const zc = ws.getCell(`A${rowNum}`);
      zc.value = `>  ${ZONE_NAMES[zk]}`;
      zc.font = { name:"Arial", bold:true, italic:true, size:9, color:{argb:"FF"+WHITE} };
      zc.fill = { type:"pattern", pattern:"solid", fgColor:{argb:"FF"+(ZONE_COLORS[zk]||"333333")} };
      zc.alignment = { horizontal:"left", vertical:"middle", indent:1 };
      ws.getRow(rowNum).height = 14;
      rowNum++;
    }

    dataN++;
    const row = ws.getRow(rowNum);
    row.height = 15;

    const sc = (col, val, align, bold, color, bgc, accent) => {
      const c = row.getCell(col);
      c.value = val;
      c.font = { name:"Arial", size:9, bold:!!bold, color:{argb:"FF"+(color||DARK)} };
      c.fill = { type:"pattern", pattern:"solid", fgColor:{argb:"FF"+(bgc||bg)} };
      c.alignment = { horizontal:align||"center", vertical:"middle" };
      c.border = accent
        ? { left:{style:"medium",color:{argb:"FF"+gc}}, right:tb().right, top:tb().top, bottom:tb().bottom }
        : tb();
    };

    sc(1,  String(dataN).padStart(2,"0"),       "center", false, "888888", "F7F7F7");
    sc(2,  String(item.grupo).padStart(2,"0"),  "center", true,  WHITE,    gc);
    sc(3,  item.insumo,       "center", false, DARK, bg, true);
    sc(4,  item.descripcion,  "left",   false, DARK, bg);
    sc(5,  item.cantidad,     "right",  true,  DARK, bg);
    sc(6,  item.um,           "center", false, "555555", bg);
    sc(7,  item.ubicacion||"—","center",true,  DARK, bg);
    sc(8,  item.producto,     "center", true,  DARK, bg);
    sc(9,  item.fecha,        "center", false, "555555", bg);
    sc(10, item.of,           "center", true,  DARK, bg);
    rowNum++;
  }

  ws.mergeCells(`A${rowNum}:J${rowNum}`);
  const pie = ws.getCell(`A${rowNum}`);
  pie.value = `Total: ${dataN} items  |  ${ordersInfo.length} ordenes  |  ${new Date().toLocaleDateString("es-AR")}`;
  pie.font = { name:"Arial", bold:true, size:9, color:{argb:"FF"+DARK} };
  pie.fill = { type:"pattern", pattern:"solid", fgColor:{argb:"FF"+GOLD} };
  pie.alignment = { horizontal:"center", vertical:"middle" };
  ws.getRow(rowNum).height = 16;

  const fecha = new Date().toLocaleDateString("es-AR").replace(/\//g,"");
  res.setHeader("Content-Type","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition",`attachment; filename="Consolidado_Picking_${fecha}.xlsx"`);
  await wb.xlsx.write(res);
  res.end();
});

app.get("*", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "index.html"));
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Picking app corriendo en puerto ${PORT}`));
