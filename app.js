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

// Espaçamentos (tweak aqui)
const SPACE_AFTER_TITLE = 120;
const SPACE_AFTER_PROCESS_LINE = 120;
const SPACE_AFTER_ORGAO = 80;
const SPACE_AFTER_TIPO = 80;
const SPACE_AFTER_INTERESSADO = 60;
const SPACE_AFTER_ADV = 50;

function setStatus(msg) {
  statusEl.textContent = msg;
}

function splitLines(value) {
  if (!value) return [];
  return String(value).split(/\r?\n/).map(v => v.trim()).filter(Boolean);
}

function upper(v) {
  return String(v ?? "").trim().toUpperCase();
}

// remove acentos e normaliza pra comparação
function normalizeName(s) {
  return String(s ?? "")
    .trim()
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
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
  btnPreview.disabled = !rows;
  btnGenerate.disabled = !(rows && headerOk());
}

updateButtons();

[sessionNumberEl, sessionTypeEl, sessionDateEl].forEach((el) => {
  if (!el) return;
  el.addEventListener("input", updateButtons);
  el.addEventListener("change", updateButtons);
});

// ===== Leitura do XLSX =====
fileInput.addEventListener("change", async (e) => {
  previewEl.textContent = "";
  rows = null;
  updateButtons();

  const file = e.target.files?.[0];
  if (!file) {
    setStatus("Nenhum arquivo selecionado.");
    return;
  }

  if (!file.name.toLowerCase().endsWith(".xlsx")) {
    setStatus("Selecione um arquivo .xlsx.");
    return;
  }

  setStatus("Lendo XLSX...");

  try {
    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    setStatus(`XLSX OK. Linhas: ${rows.length}.`);
    updateButtons();
  } catch (err) {
    console.error(err);
    setStatus("Erro ao ler XLSX. Abra o Console (F12) e veja o erro.");
    rows = null;
    updateButtons();
  }
});

// ===== Prévia =====
btnPreview.addEventListener("click", () => {
  if (!rows) return;
  previewEl.textContent = JSON.stringify(rows.slice(0, 10), null, 2);
});

// ===== Regras de tipo do relator =====
const CONSELHEIROS = [
  "VALDECIR PASCOAL",
  "RANILSON RAMOS",
  "DIRCEU RODOLFO DE MELO JUNIOR",
  "MARCOS LORETO",
  "CARLOS NEVES",
  "EDUARDO LYRA PORTO",
  "RODRIGO NOVAES",
].map(normalizeName);

function relatorPrefix(relatorRaw) {
  const n = normalizeName(relatorRaw);

  // heurística: se o nome do relator "contém" o nome-chave, marca como conselheiro
  const isConselheiro = CONSELHEIROS.some((key) => n.includes(key));
  return isConselheiro ? "CONSELHEIRO" : "CONSELHEIRO SUBSTITUTO";
}

// ===== DOCX =====
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

    const separator = () =>
      new Paragraph({
        children: [new TextRun("______________________________________________________________________________________")],
        spacing: { before: 0, after: 0 },
      });

    const blankLine = (after = 80) =>
      new Paragraph({
        children: [new TextRun(" ")],
        spacing: { after },
      });

    // ===== Cabeçalho centralizado =====
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
        spacing: { after: 120 },
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
        spacing: { after: 80 },
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
        spacing: { after: 140 },
      })
    );

    children.push(separator());

    // Agrupa por relator (mantém ordem de aparição)
    const grouped = new Map();
    for (const r of rows) {
      const rel = String(r["Relator"] ?? "").trim();
      if (!grouped.has(rel)) grouped.set(rel, []);
      grouped.get(rel).push(r);
    }

    for (const [relator, processos] of grouped.entries()) {
      const prefix = relatorPrefix(relator);

      // RELATOR: em negrito, tamanho 11
      children.push(
        new Paragraph({
          children: [
            new TextRun({
              text: `RELATOR: ${prefix} ${upper(relator)}`,
              bold: true,
              size: SIZE_HEADER,
              font: FONT,
            }),
          ],
          spacing: { before: 240, after: 0 },
        })
      );

      // (2) Linha em branco entre relator e primeiro processo
      children.push(blankLine(120));

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

        // Linha do processo (negrito). Só o label colorido.
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
            spacing: { after: SPACE_AFTER_PROCESS_LINE },
          })
        );

        // (1) Órgão: AGORA EM NEGRITO
        children.push(
          new Paragraph({
            children: [
              new TextRun({
                text: upper(row["Órgão"]),
                bold: true,
                size: SIZE_BODY,
                font: FONT,
              }),
            ],
            spacing: { after: SPACE_AFTER_ORGAO },
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
            spacing: { after: SPACE_AFTER_TIPO },
          })
        );

        // Interessados: sem negrito
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
              spacing: { after: SPACE_AFTER_INTERESSADO },
            })
          );
        });

        // Advogados: sem negrito
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
              spacing: { after: SPACE_AFTER_ADV },
            })
          );
        });

        // Espaço entre processos (um respiro)
        children.push(blankLine(120));
      }

      children.push(separator());
    }

    const doc = new Document({ sections: [{ children }] });
    const blob = await Packer.toBlob(doc);

    const filename = `pauta_${dateBR.replaceAll("/", "-")}.docx`;
    saveAs(blob, filename);

    setStatus(`DOCX gerado: ${filename}`);
  } catch (err) {
    console.error(err);
    setStatus("Erro ao gerar DOCX. Abra o Console (F12) e veja o erro.");
  }
});
