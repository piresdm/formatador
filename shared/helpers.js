export function clearElement(el) {
  if (!el) return;
  el.innerHTML = "";
}

export function splitLines(value) {
  if (!value) return [];
  return String(value)
    .split(/\r?\n/)
    .map((v) => v.trim())
    .filter(Boolean);
}

export function upper(v) {
  return String(v ?? "").trim().toUpperCase();
}

// remove acentos e normaliza pra comparação
export function normalizeName(s) {
  return String(s ?? "")
    .trim()
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

export function formatDateBR(date) {
  if (!date) return "";
  const [y, m, d] = String(date).split("-");
  if (!y || !m || !d) return "";
  return `${d}/${m}/${y}`;
}

export function ordinalFeminino(n) {
  const num = Number(n);
  if (!Number.isFinite(num) || num < 1) return "";
  return `${Math.trunc(num)}ª`;
}

/**
 * Lê a primeira aba do XLSX e devolve array de objetos.
 * Depende do XLSX global (SheetJS via CDN), que fica em window.XLSX.
 */
export async function readFirstSheetXlsxToJson(file) {
  if (!file) throw new Error("Arquivo ausente.");
  if (!file.name?.toLowerCase?.().endsWith(".xlsx")) {
    throw new Error("Selecione um arquivo .xlsx.");
  }
  if (!window.XLSX) {
    throw new Error("Biblioteca XLSX não carregada (CDN).");
  }

  const arrayBuffer = await file.arrayBuffer();
  const workbook = window.XLSX.read(arrayBuffer, { type: "array" });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  return window.XLSX.utils.sheet_to_json(sheet, { defval: "" });
}

export function safeFilename(s) {
  return String(s ?? "")
    .trim()
    .replace(/[\\/:*?"<>|]+/g, "_")
    .replace(/\s+/g, "_");
}
