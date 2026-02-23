/* global XLSX, docx, saveAs */

const fileInput = document.getElementById("fileInput");
const btnPreview = document.getElementById("btnPreview");
const btnGenerate = document.getElementById("btnGenerate");
const statusEl = document.getElementById("status");
const previewEl = document.getElementById("preview");

// NOVO: campos do cabeçalho
const sessionTypeEl = document.getElementById("sessionType");
const sessionDateEl = document.getElementById("sessionDate");
const sessionTimeEl = document.getElementById("sessionTime");

let rows = null;

function setStatus(msg) {
  statusEl.textContent = msg;
}

function clearPreview() {
  previewEl.textContent = "";
}

function splitLines(value) {
  if (!value) return [];
  const s = String(value);
  return s
    .split(/\r?\n/)
    .map((x) => x.trim())
    .filter(Boolean);
}

function upper(value) {
  return String(value ?? "").trim().toUpperCase();
}

function groupBy(arr, key) {
  const m = new Map();
  for (const item of arr) {
    const k = String(item[key] ?? "").trim();
    if (!m.has(k)) m.set(k, []);
    m.get(k).push(item);
  }
  return m;
}

function formatDateBR(yyyyMmDd) {
  if (!yyyyMmDd) return "";
  const [y, m, d] = String(yyyyMmDd).split("-");
  if (!y || !m || !d) return "";
  return `${d}/${m}/${y}`;
}

// ===== Leitura do XLSX =====
fileInput.addEventListener("change", async (e) => {
  clearPreview();
  rows = null;

  const file = e.target.files?.[0];
  if (!file) {
    btnPreview.disabled = true;
    btnGenerate.disabled = true;
    setStatus("Nenhum arquivo selecionado.");
    return;
  }

  if (!file.name.toLowerCase().endsWith(".xlsx")) {
    btnPreview.disabled = true;
    btnGenerate.disabled = true;
    setStatus("Selecione um arquivo .xlsx.");
    return;
  }

  setStatus("Lendo arquivo...");
  try {
    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });

    const firstSheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[firstSheetName];

    rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    btnPreview.disabled = false;
    btnGenerate.disabled = false;

    setStatus(`OK: ${file.name} | Aba: "${firstSheetName}" | Linhas: ${rows.length}`);
  } catch (err) {
    console.error(err);
    btnPreview.disabled = true;
    btnGenerate.disabled = true;
    setStatus("Erro ao ler o arquivo. Veja o console do navegador.");
  }
});

btnPreview.addEventListener("click", () => {
  if (!rows) return;
  const sample = rows.slice(0, 10);
  previewEl.textContent = JSON.stringify(sample, null, 2);
});

// ===== Geração do DOCX =====
btnGenerate.addEventListener("click", async () => {
  if (!rows) return;

  setStatus("Gerando DOCX...");

  try {
    const { Document, Packer, Paragraph, TextRun, AlignmentType } = docx;

    const SEPARATOR =
      "______________________________________________________________________________________";

    const makeSeparator = () =>
      new Paragraph({
        children: [new TextRun(SEPARATOR)],
        spacing: { before: 180, after: 180 },
      });

    const makeRelatorHeader = (relator) =>
      new Paragraph({
        children: [
          new TextRun({
            text: `RELATOR: ${upper(relator)}`,
            bold: true,
          }),
        ],
        spacing: { before: 220, after: 120 },
      });

    const makeProcessTitle = (row) => {
      const sistema = upper(row["Sistema de Tramitação"]);
      const proc = String(row["Processo"] ?? "").trim();
      const voto = upper(row["Voto"]);

      let label = "PROCESSO";
      let color = "000000";

      if (sistema === "E-TCE") {
        label = "PROCESSO ELETRÔNICO eTCE";
        color = "FF0000"; // vermelho
      } else if (sistema === "AP") {
        label = "PROCESSO DIGITAL TCE";
        color = "0070C0"; // azul
      }

      const suffix = voto === "LISTADO" ? " (Voto em lista)" : "";

      return new Paragraph({
        children: [
          new TextRun({
            text: `${label} Nº ${proc}${suffix}`,
            bold: true,
            color,
          }),
        ],
        spacing: { before: 140, after: 80 },
      });
    };

    const makeUpperLine = (text) =>
      new Paragraph({
        children: [new TextRun({ text: upper(text) })],
        spacing: { after: 40 },
      });

    const makePlainLine = (text) =>
      new Paragraph({
        children: [new TextRun({ text: String(text ?? "").trim() })],
        spacing: { after: 40 },
      });

    const makeAdvLine = (lawyer) =>
      new Paragraph({
        children: [new TextRun({ text: `(Adv. ${lawyer})` })],
        spacing: { after: 20 },
      });

    // ===== Cabeçalho vindo da UI =====
    const sessionType = sessionTypeEl?.value || "PLENO";
    const sessionDate = formatDateBR(sessionDateEl?.value);

    const children = [];

    children.push(
      new Paragraph({
        children: [
          new TextRun({
            text: `PAUTA DA SESSÃO ORDINÁRIA DO ${sessionType}`,
            bold: true,
            size: 28,
          }),
        ],
        alignment: AlignmentType.CENTER,
        spacing: { after: 120 },
      })
    );

    children.push(
      new Paragraph({
        children: [new TextRun({ text: `DATA: ${sessionDate}`, bold: true })],
        spacing: { after: 40 },
      })
    );

      children.push(
    new Paragraph({
      children: [new TextRun({ text: `HORÁRIO: 10h`, bold: true })],
      spacing: { after: 120 },
    })
  );

    children.push(makeSeparator());

    // ===== Agrupar por relator =====
    const cleaned = rows
      .map((r) => ({ ...r, Relator: String(r["Relator"] ?? "").trim() }))
      .filter((r) => r.Relator);

    const byRelator = groupBy(cleaned, "Relator");

    for (const [relator, items] of byRelator.entries()) {
      children.push(makeRelatorHeader(relator));
      children.push(new Paragraph({ text: "", spacing: { after: 60 } }));

      for (const row of items) {
        children.push(makeProcessTitle(row));

        // Órgão (normalmente já vem com "- ano")
        children.push(makeUpperLine(row["Órgão"]));

        // Tipo Processo
        children.push(makeUpperLine(row["Tipo Processo"]));

        // Interessados (1 por linha)
        const interessados = splitLines(row["Interessados"]);
        for (const it of interessados) children.push(makePlainLine(it));

        // Advogados (1 por linha, com prefixo Adv.)
        const advs = splitLines(row["Advogados"]);
        for (const adv of advs) children.push(makeAdvLine(adv));

        // Espaço entre processos
        children.push(new Paragraph({ text: "", spacing: { after: 140 } }));
      }

      children.push(makeSeparator());
    }

    const doc = new Document({
      sections: [{ children }],
    });

    const blob = await Packer.toBlob(doc);

    const filename = `pauta_${new Date().toISOString().slice(0, 10)}.docx`;
    saveAs(blob, filename);

    setStatus(`DOCX gerado: ${filename}`);
  } catch (err) {
    console.error(err);
    setStatus("Erro ao gerar DOCX. Veja o console do navegador.");
  }
});
