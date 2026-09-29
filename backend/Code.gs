/**
 * BACKEND - Mis Finanzas (v2)
 * Va pegado en Extensiones > Apps Script de tu Google Sheet.
 *
 * Pestaña "Gastos" (fila 1 = encabezados):
 *   A Timestamp | B Fecha | C Monto | D Categoria | E MetodoPago | F Nota | G Fuente | H Comercio | I Id
 * Pestaña "Reglas" (se crea sola):
 *   A Comercio (en minúsculas) | B Categoria
 *   Puedes editarla a mano: si pones "wong", cualquier comercio que contenga "wong" va a esa categoría.
 *
 * Qué recibe (POST, cuerpo JSON):
 *   monto       número o texto ("12.5", "S/ 12.50", "PEN 1,250.00")
 *   categoria   opcional. Si falta y hay comercio, se busca en Reglas; si no hay regla -> "Sin categorizar"
 *   metodoPago  "Efectivo" | "Yape/Plin" | "Tarjeta" | "Apple Pay" ...
 *   nota, comercio, tarjeta, fuente ("app" | "atajo" | "wallet"), fecha (opcional)
 *
 * GET ?action=list                          -> todos los gastos
 * GET ?action=delete&id=...                 -> borra un gasto
 * GET ?action=categorize&id=...&categoria=  -> cambia la categoría y aprende la regla del comercio
 */

const SHEET_NAME = "Gastos";
const RULES_SHEET = "Reglas";
const HEADERS = ["Timestamp", "Fecha", "Monto", "Categoria", "MetodoPago", "Nota", "Fuente", "Comercio", "Id"];
const SIN_CATEGORIA = "Sin categorizar";

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getRulesSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(RULES_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(RULES_SHEET);
    sheet.appendRow(["Comercio", "Categoria"]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/* ---------- helpers ---------- */

function parseMonto_(v) {
  if (typeof v === "number") return Math.abs(v);
  let s = String(v || "").replace(/[^\d.,-]/g, "").replace(/^[.,-]+/, "");
  if (s.indexOf(",") > -1 && s.indexOf(".") > -1) {
    // el último separador es el decimal
    if (s.lastIndexOf(",") > s.lastIndexOf(".")) s = s.replace(/\./g, "").replace(",", ".");
    else s = s.replace(/,/g, "");
  } else if (s.indexOf(",") > -1) {
    const dec = s.split(",").pop();
    s = dec.length === 2 ? s.replace(",", ".") : s.replace(/,/g, "");
  }
  const n = Math.abs(parseFloat(s));
  return isNaN(n) ? 0 : Math.round(n * 100) / 100;
}

function normalize_(s) {
  return String(s || "")
    .toLowerCase()
    .normalize("NFD").replace(/[̀-ͯ]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function findRule_(comercio) {
  const key = normalize_(comercio);
  if (!key) return null;
  const rows = getRulesSheet_().getDataRange().getValues().slice(1);
  let best = null;
  rows.forEach(function (r) {
    const k = normalize_(r[0]);
    if (!k || !r[1]) return;
    if (key === k || key.indexOf(k) > -1) {
      if (!best || k.length > best.key.length) best = { key: k, categoria: String(r[1]) };
    }
  });
  return best ? best.categoria : null;
}

function learnRule_(comercio, categoria) {
  const key = normalize_(comercio);
  if (!key || !categoria || categoria === SIN_CATEGORIA) return;
  const sheet = getRulesSheet_();
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (normalize_(values[i][0]) === key) {
      sheet.getRange(i + 1, 2).setValue(categoria);
      return;
    }
  }
  sheet.appendRow([key, categoria]);
}

/* ---------- endpoints ---------- */

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const data = JSON.parse((e.postData && e.postData.contents) || "{}");
    const sheet = getSheet_();

    const monto = parseMonto_(data.monto);
    if (!(monto > 0)) return jsonResponse_({ ok: false, error: "monto inválido" });

    const comercio = String(data.comercio || "").trim();
    let categoria = String(data.categoria || "").trim();
    if (!categoria) categoria = (comercio && findRule_(comercio)) || SIN_CATEGORIA;

    let metodo = String(data.metodoPago || "").trim();
    if (data.tarjeta) metodo = (metodo || "Apple Pay") + " · " + String(data.tarjeta).trim();

    const id = Utilities.getUuid();
    const now = new Date();
    const fecha = data.fecha ? new Date(data.fecha) : now;

    sheet.appendRow([
      now,
      isNaN(fecha.getTime()) ? now : fecha,
      monto,
      categoria,
      metodo,
      data.nota || "",
      data.fuente || "app",
      comercio,
      id
    ]);

    return jsonResponse_({
      ok: true,
      id: id,
      monto: monto,
      categoria: categoria,
      mensaje: "S/ " + monto.toFixed(2) + " · " + categoria + (comercio ? " · " + comercio : "")
    });
  } catch (err) {
    return jsonResponse_({ ok: false, error: err.message });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  try {
    const action = (e.parameter && e.parameter.action) || "list";

    if (action === "list") {
      const values = getSheet_().getDataRange().getValues();
      const headers = values.shift();
      const rows = values.map(function (row) {
        const obj = {};
        headers.forEach(function (h, i) {
          obj[h] = row[i] instanceof Date ? row[i].toISOString() : row[i];
        });
        return obj;
      });
      return jsonResponse_({ ok: true, data: rows });
    }

    if (action === "delete" || action === "categorize") {
      const id = e.parameter.id;
      const sheet = getSheet_();
      const values = sheet.getDataRange().getValues();
      for (let i = 1; i < values.length; i++) {
        if (values[i][8] === id) {
          if (action === "delete") {
            sheet.deleteRow(i + 1);
          } else {
            const categoria = String(e.parameter.categoria || "").trim();
            sheet.getRange(i + 1, 4).setValue(categoria);
            learnRule_(values[i][7], categoria);
          }
          return jsonResponse_({ ok: true });
        }
      }
      return jsonResponse_({ ok: false, error: "no encontrado" });
    }

    return jsonResponse_({ ok: false, error: "acción desconocida" });
  } catch (err) {
    return jsonResponse_({ ok: false, error: err.message });
  }
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
