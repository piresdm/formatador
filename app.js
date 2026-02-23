/* global XLSX, docx, saveAs */

const fileInput = document.getElementById("fileInput");
const btnPreview = document.getElementById("btnPreview");
const btnGenerate = document.getElementById("btnGenerate");
const statusEl = document.getElementById("status");
const previewEl = document.getElementById("preview");

// Campos de cabeçalho (obrigatórios)
const sessionNumberEl = document.getElementById("sessionNumber");
const sessionTypeEl = document.getElementById("sessionType");
const sessionDateEl = document.getElementById("sessionDate");

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

// 1 -> 1ª, 2 -> 2ª, 3 -> 3ª ...
function ordinalFeminino(n) {
  const num = Number(n);
  if (!Number.isFinite(num) || num < 1) return "";
  return `${Math.trunc(num)}ª`;
}

function getHeaderValues() {
  const sessionNumberRaw = String(sessionNumberEl?.value ?? "").trim();
  const sessionType = String(sessionTypeEl?.value ?? "").trim();
  const sessionDate = String(sessionDateEl?.value ?? "").trim();

  const sessionNumberOrd = ordinalFeminino(sessionNumberRaw);
  const dateBR = formatDateBR(sessionDate);

  return { sessionNumberOrd, sessionType, dateBR };
}

function canGenerate() {
  const { sessionNumberOrd, sessionType, dateBR } = getHeaderValues();
  const hasFileRows = Array.isArray(rows) && rows.length >= 0; // permite vazio, mas tem que ter lido
  const headerOk = !!sessionNumberOrd && !!sessionType && !!dateBR;
  return hasFileRows && headerOk;
}

function updateButtons() {
  const headerOk = !!getHeaderValues().sessionNumberOrd && !!getHeaderValues().sessionType && !!getHeaderValues().dateBR;

  btnPreview.disabled = !rows; // preview só depende do xlsx lido
  btnGenerate.disabled = !(rows && headerOk);

  if (!rows) {
    setStatus("Nenhum arquivo selecionado.");
    return;
  }

  if (!headerOk) {
    setStatus("Preencha Nº da sessão, Tipo de sessão e Data para liberar a geração.");
    return;
  }

  setStatus("Pronto para gerar.");
}

// Reagir a mudanças nos campos do cabeçalho
[sessionNumberEl, sessionTypeEl, sessionDateEl].forEach((el) => {
  if (!el) return;
  el.addEventListener("input", updateButtons);
  el.addEventListener("change", updateButtons);
});

// ===== Leitura do XLSX =====
fileInput.addEventListener("change", async (e) => {
  clearPreview();
  rows = null;

  const file = e.target.files?.[0];
  if (!file) {
    updateButtons();
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

    setStatus(`OK: ${file.name} | Aba: "${firstSheetName}" | Linhas: ${rows.length}`);
    updateButtons();
  } catch (err) {
    console.error(err);
    rows = null;
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

  const { sessionNumberOrd, sessionType, dateBR } = getHeaderValues();
  if (!sessionNumberOrd || !sessionType || !dateBR) {
    setStatus("Preencha Nº da sessão, Tipo de sessão e Data.");
    return;
  }

  setStatus("Gerando DOCX...");

  try {
    const { Document, Packer, Paragraph, TextRun, AlignmentType } = docx;

    const SEPARATOR =
      "______________________________________________________________________________________";

    const makeSeparator = () =>
      new Paragraph({
        children: [new TextRun(SEPARATOR)],
        spacing: { before: 0, after: 0 },
      });

    // No seu modelo: RELATOR em negrito (linha toda) :contentReference[oaicite:1]{index=1}
    const makeRelatorHeader = (relator) =>
      new Paragraph({
        children: [
          new TextRun({
            text: `RELATOR: ${upper(relator)}`,
            bold: true,
          }),
        ],
        spacing: { before: 240, after: 120 },
      });

    // Linha do processo:
    // - SOMENTE o rótulo colorido
    // - "Nº {proc} (Voto em lista)" preto
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
          new TextRun({ text: `${label} `, bold: true, color }), // colorido
          new TextRun({ text: `Nº ${proc}${suffix}`, bold: true, color: "000000" }), // preto
        ],
        spacing: { before: 120, after: 80 },
      });
    };

    // Órgão e Tipo Processo aparecem em caixa alta, sem negrito no modelo :contentReference[oaicite:2]{index=2}
    const makeUpperLine = (text) =>
      new Paragraph({
        children: [new TextRun({ text: upper(text) })],
        spacing: { after: 40 },
      });

    // Interessados sem parênteses, 1 por linha
    const makePlainLine = (text) =>
      new Paragraph({
        children: [new TextRun({ text: String(text ?? "").trim() })],
        spacing: { after: 20 },
      });

    // Advogados com (Adv. ...)
    const makeAdvLine = (lawyer) =>
      new Paragraph({
        children: [new TextRun({ text: `(Adv. ${lawyer})` })],
        spacing: { after: 10 },
      });

    // ===== Cabeçalho igual ao modelo (com ordinal e horário fixo) :contentReference[oaicite:3]{index=3} =====
    const children = [];

    children.push(
      new Paragraph({
        children: [
          new TextRun({
            text: `PAUTA DA ${sessionNumberOrd} SESSÃO ORDINÁRIA DO ${upper(sessionType)}`,
            bold: true,
            size: 28,
          }),
        ],
        spacing: { after: 80 },
      })
    );

    children.push(
      new Paragraph({
        children: [new TextRun({ text: `DATA: ${dateBR}`, bold: true })],
        spacing: { after: 40 },
      })
    );

    children.push(
      new Paragraph({
        children: [new TextRun({ text: `HORÁRIO: 10h`, bold: true })],
        spacing: { after: 60 },
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
      children.push(new Paragraph({ text: "" }));

      for (const row of items) {
        children.push(makeProcessTitle(row));

        // Órgão
        children.push(makeUpperLine(row["Órgão"]));

        // Tipo Processo
        children.push(makeUpperLine(row["Tipo Processo"]));

        // Interessados (1 por linha)
        const interessados = splitLines(row["Interessados"]);
        for (const it of interessados) children.push(makePlainLine(it));

        // Advogados (1 por linha)
        const advs = splitLines(row["Advogados"]);
        for (const adv of advs) children.push(makeAdvLine(adv));

        // Espaço entre processos
        children.push(new Paragraph({ text: "", spacing: { after: 120 } }));
      }

      children.push(makeSeparator());
    }

    const doc = new Document({
      sections: [{ children }],
    });

    const blob = await Packer.toBlob(doc);

    // Nome do arquivo com data escolhida
    const filename = `pauta_${dateBR.replaceAll("/", "-")}.docx`;
    saveAs(blob, filename);

    setStatus(`DOCX gerado: ${filename}`);
  } catch (err) {
    console.error(err);
    setStatus("Erro ao gerar DOCX. Veja o console do navegador.");
  }
});

// Estado inicial
updateButtons();
