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
    const k = String(item[key] ?? "").trim();
    if (!map.has(k)) map.set(k, []);
    map.get(k).push(item);
  }
  return map;
}

function formatDateBR(date) {
  if (!date) return "";
  const [y, m, d] = String(date).split("-");
  if (!y || !m || !d) return "";
  return `${d}/${m}/${y}`;
}

function ordinalFeminino(n) {
  const num = Number(n);
  if (!Number.isFinite(num) || num < 1) return "";
  return `${Math.trunc(num)}ª`;
}

function headerOk() {
  return (
    !!ordinalFeminino(sessionNumberEl?.value) &&
    !!String(sessionTypeEl?.value || "").trim() &&
    !!String(sessionDateEl?.value || "").trim()
  );
}

function updateButtons() {
  // Preview só depende de ter lido o arquivo
  btnPreview.disabled = !rows;

  // Gerar depende do arquivo + cabeçalho preenchido
  btnGenerate.disabled = !(rows && headerOk());

  if (!rows) {
    setStatus("Nenhum arquivo selecionado.");
    return;
  }
  if (!headerOk()) {
    setStatus("Arquivo carregado. Preencha Nº da sessão, Tipo de sessão e Data para liberar a geração.");
    return;
  }
  setStatus("Pronto para gerar.");
}

// Listeners para reavaliar os botões
[fileInput, sessionNumberEl, sessionTypeEl, sessionDateEl].forEach((el) => {
  if (!el) return;
  el.addEventListener("input", updateButtons);
  el.addEventListener("change", updateButtons);
});

// ===== Leitura do XLSX =====
fileInput.addEventListener("change", async (e) => {
  previewEl.textContent = "";
  rows = null;

  const file = e.target.files?.[0];
  if (!file) {
    updateButtons();
    return;
  }

  if (!file.name.toLowerCase().endsWith(".xlsx")) {
    rows = null;
    btnPreview.disabled = true;
    btnGenerate.disabled = true;
    setStatus("Selecione um arquivo .xlsx.");
    return;
  }

  setStatus("Lendo arquivo...");

  try {
    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    setStatus(`Arquivo carregado. Linhas: ${rows.length}`);
    updateButtons();
  } catch (err) {
    console.error(err);
    rows = null;
    setStatus("Erro ao ler o arquivo. Veja o console do navegador.");
    updateButtons();
  }
});

// ===== Pré-visualização =====
btnPreview.addEventListener("click", () => {
  if (!rows) return;
  const sample = rows.slice(0, 10);
  previewEl.textContent = JSON.stringify(sample, null, 2);
});

// ===== Geração do DOCX =====
btnGenerate.addEventListener("click", async () => {
  if (!rows) return;
  if (!headerOk()) {
    setStatus("Preencha Nº da sessão, Tipo de sessão e Data.");
    return;
  }

  setStatus("Gerando DOCX...");

  try {
    const { Document, Packer, Paragraph, TextRun, AlignmentType } = docx;

    const sessionNumber = ordinalFeminino(sessionNumberEl.value);
    const sessionType = upper(sessionTypeEl.value);
    const dateBR = formatDateBR(sessionDateEl.value);

    const children = [];

    // Cabeçalho centralizado (igual ao modelo)
    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({
            text: `PAUTA DA ${sessionNumber} SESSÃO ORDINÁRIA DO ${sessionType}`,
            bold: true,
            size: SIZE_HEADER,
            font: FONT,
          }),
        ],
        spacing: { after: 80 },
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
            font: FONT,
          }),
        ],
        spacing: { after: 40 },
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
            font: FONT,
          }),
        ],
        spacing: { after: 100 },
      })
    );

    children.push(
      new Paragraph({
        children: [new TextRun("______________________________________________________________________________________")],
      })
    );

    // Agrupa por relator
    const grouped = groupBy(rows, "Relator");

    for (const [relator, processos] of grouped.entries()) {
      // Relator: negrito, tamanho 11
      children.push(
        new Paragraph({
          children: [
            new TextRun({
              text: `RELATOR: ${upper(relator)}`,
              bold: true,
              size: SIZE_HEADER,
              font: FONT,
            }),
          ],
          spacing: { before: 200, after: 100 },
        })
      );

      for (const row of processos) {
        const sistema = upper(row["Sistema de Tramitação"]);
        const processo = String(row["Processo"] ?? "").trim();
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

        // Linha do processo:
        // - label colorido
        // - Nº + número + (Voto em lista) preto
        // - tudo em negrito (porque não é órgão/interessado/adv)
        children.push(
          new Paragraph({
            children: [
              new TextRun({
                text: `${label} `,
                bold: true,
                color,
                size: SIZE_BODY,
                font: FONT,
              }),
              new TextRun({
                text: `Nº ${processo}${voto}`,
                bold: true,
                color: "000000",
                size: SIZE_BODY,
                font: FONT,
              }),
            ],
            spacing: { after: 60 },
          })
        );

        // Órgão: SEM negrito
        children.push(
          new Paragraph({
            children: [
              new TextRun({
                text: upper(row["Órgão"]),
                bold: false,
                size: SIZE_BODY,
                font: FONT,
              }),
            ],
          })
        );

        // Tipo Processo: negrito
        children.push(
          new Paragraph({
            children: [
              new TextRun({
                text: upper(row["Tipo Processo"]),
                bold: true,
                size: SIZE_BODY,
                font: FONT,
              }),
            ],
          })
        );

        // Interessados: SEM negrito
        splitLines(row["Interessados"]).forEach((i) => {
          children.push(
            new Paragraph({
              children: [
                new TextRun({
                  text: i,
                  bold: false,
                  size: SIZE_BODY,
                  font: FONT,
                }),
              ],
            })
          );
        });

        // Advogados: SEM negrito
        splitLines(row["Advogados"]).forEach((a) => {
          children.push(
            new Paragraph({
              children: [
                new TextRun({
                  text: `(Adv. ${a})`,
                  bold: false,
                  size: SIZE_BODY,
                  font: FONT,
                }),
              ],
            })
          );
        });

        children.push(new Paragraph({ text: "" }));
      }

      children.push(
        new Paragraph({
          children: [new TextRun("______________________________________________________________________________________")],
        })
      );
    }

    const doc = new Document({ sections: [{ children }] });
    const blob = await Packer.toBlob(doc);

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
