/* global XLSX, docx, saveAs */

const fileInput = document.getElementById("fileInput");
const btnPreview = document.getElementById("btnPreview");
const btnGenerate = document.getElementById("btnGenerate");
const statusEl = document.getElementById("status");
const previewEl = document.getElementById("preview");

const sessionNumberEl = document.getElementById("sessionNumber");
const sessionTypeEl = document.getElementById("sessionType");
const sessionDateEl = document.getElementById("sessionDate");

let rows = null;

const FONT = "Roboto";
const SIZE_HEADER = 22; // 11pt
const SIZE_BODY = 20;   // 10pt

function setStatus(msg) {
  statusEl.textContent = msg;
}

function splitLines(value) {
  if (!value) return [];
  return String(value)
    .split(/\r?\n/)
    .map(v => v.trim())
    .filter(Boolean);
}

function upper(v) {
  return String(v ?? "").trim().toUpperCase();
}

function groupBy(arr, key) {
  const map = new Map();
  for (const item of arr) {
    const k = item[key];
    if (!map.has(k)) map.set(k, []);
    map.get(k).push(item);
  }
  return map;
}

function formatDateBR(date) {
  if (!date) return "";
  const [y, m, d] = date.split("-");
  return `${d}/${m}/${y}`;
}

function ordinalFeminino(n) {
  const num = Number(n);
  if (!num) return "";
  return `${num}ª`;
}

function updateButtons() {
  const ok =
    rows &&
    sessionNumberEl.value &&
    sessionTypeEl.value &&
    sessionDateEl.value;

  btnGenerate.disabled = !ok;
}

[fileInput, sessionNumberEl, sessionTypeEl, sessionDateEl]
  .forEach(el => el && el.addEventListener("change", updateButtons));

// ================= XLSX =================

fileInput.addEventListener("change", async (e) => {
  const file = e.target.files?.[0];
  if (!file) return;

  const arrayBuffer = await file.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

  setStatus("Arquivo carregado.");
  updateButtons();
});

// ================= DOCX =================

btnGenerate.addEventListener("click", async () => {

  const { Document, Packer, Paragraph, TextRun, AlignmentType } = docx;

  const children = [];

  const sessionNumber = ordinalFeminino(sessionNumberEl.value);
  const sessionType = upper(sessionTypeEl.value);
  const dateBR = formatDateBR(sessionDateEl.value);

  // ===== Cabeçalho centralizado =====

  children.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun({
          text: `PAUTA DA ${sessionNumber} SESSÃO ORDINÁRIA DO ${sessionType}`,
          bold: true,
          size: SIZE_HEADER,
          font: FONT
        })
      ],
      spacing: { after: 80 }
    })
  );

  children.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun({
          text: `DATA: ${dateBR}`,
          bold: true,
          size: SIZE_HEADER,
          font: FONT
        })
      ],
      spacing: { after: 40 }
    })
  );

  children.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun({
          text: `HORÁRIO: 10h`,
          bold: true,
          size: SIZE_HEADER,
          font: FONT
        })
      ],
      spacing: { after: 100 }
    })
  );

  children.push(new Paragraph({ text: "______________________________________________________________________________________" }));

  const grouped = groupBy(rows, "Relator");

  for (const [relator, processos] of grouped.entries()) {

    children.push(
      new Paragraph({
        children: [
          new TextRun({
            text: `RELATOR: ${upper(relator)}`,
            bold: true,
            size: SIZE_HEADER,
            font: FONT
          })
        ],
        spacing: { before: 200, after: 100 }
      })
    );

    for (const row of processos) {

      const sistema = upper(row["Sistema de Tramitação"]);
      const processo = row["Processo"];
      const voto = upper(row["Voto"]) === "LISTADO" ? " (Voto em lista)" : "";

      let label = "PROCESSO";
      let color = "000000";

      if (sistema === "E-TCE") {
        label = "PROCESSO ELETRÔNICO eTCE";
        color = "FF0000";
      } else if (sistema === "AP") {
        label = "PROCESSO DIGITAL TCE";
        color = "0070C0";
      }

      // Linha do processo (negrito, mas só label colorido)
      children.push(
        new Paragraph({
          children: [
            new TextRun({
              text: `${label} `,
              bold: true,
              color,
              size: SIZE_BODY,
              font: FONT
            }),
            new TextRun({
              text: `Nº ${processo}${voto}`,
              bold: true,
              color: "000000",
              size: SIZE_BODY,
              font: FONT
            })
          ],
          spacing: { after: 60 }
        })
      );

      // Órgão (sem negrito)
      children.push(
        new Paragraph({
          children: [
            new TextRun({
              text: upper(row["Órgão"]),
              size: SIZE_BODY,
              font: FONT
            })
          ]
        })
      );

      // Tipo Processo (negrito)
      children.push(
        new Paragraph({
          children: [
            new TextRun({
              text: upper(row["Tipo Processo"]),
              bold: true,
              size: SIZE_BODY,
              font: FONT
            })
          ]
        })
      );

      // Interessados (sem negrito)
      splitLines(row["Interessados"]).forEach(i => {
        children.push(
          new Paragraph({
            children: [new TextRun({ text: i, size: SIZE_BODY, font: FONT })]
          })
        );
      });

      // Advogados (sem negrito)
      splitLines(row["Advogados"]).forEach(a => {
        children.push(
          new Paragraph({
            children: [new TextRun({ text: `(Adv. ${a})`, size: SIZE_BODY, font: FONT })]
          })
        );
      });

      children.push(new Paragraph({ text: "" }));
    }

    children.push(new Paragraph({ text: "______________________________________________________________________________________" }));
  }

  const doc = new Document({
    sections: [{ children }]
  });

  const blob = await Packer.toBlob(doc);
  saveAs(blob, "pauta.docx");
});
