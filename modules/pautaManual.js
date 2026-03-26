/* global docx, saveAs */

import {
  readFirstSheetXlsxToJson,
  readPdfToText,
  readPdfToStructuredLines,
  splitLines,
  upper,
  normalizeName,
  formatDateBR,
  ordinalFeminino,
} from "../shared/helpers.js";

export function mount(container) {
  container.innerHTML = `
    <div class="module-card">
      <div class="card">
        <div class="card-body">
          <div class="row g-3 mb-3">
            <div class="col-md-3">
              <label for="sessionNumber" class="form-label">Nº da sessão</label>
              <input
                id="sessionNumber"
                type="number"
                min="1"
                step="1"
                class="form-control"
                placeholder="Ex: 1"
                required
              />
            </div>

            <div class="col-md-5">
              <label for="sessionType" class="form-label">Tipo de sessão</label>
              <select id="sessionType" class="form-select" required>
                <option value="" selected>Selecione...</option>
                <option value="PLENO">Pleno</option>
                <option value="PRIMEIRA CÂMARA">Primeira Câmara</option>
                <option value="SEGUNDA CÂMARA">Segunda Câmara</option>
              </select>
            </div>

            <div class="col-md-4">
              <label for="sessionDate" class="form-label">Data</label>
              <input id="sessionDate" type="date" class="form-control" required />
            </div>
          </div>

          <div class="row g-3">
            <div class="col-md-3">
              <label for="inputType" class="form-label">Tipo de documento</label>
              <select id="inputType" class="form-select" required>
                <option value="XLS" selected>XLS</option>
                <option value="PDF">PDF</option>
              </select>
            </div>
            <div class="col-md-9">
              <label for="fileInput" class="form-label">Selecione o arquivo</label>
              <input class="form-control" type="file" id="fileInput" accept=".xlsx" />
            </div>
          </div>

          <div class="d-flex gap-2 mt-3">
            <button id="btnPreview" class="btn btn-outline-secondary" disabled>
              Pré-visualizar (10 linhas)
            </button>

            <button id="btnGenerate" class="btn btn-primary" disabled>
              Gerar DOCX
            </button>
          </div>

          <div class="mt-3">
            <div id="status" class="small text-muted">Nenhum arquivo selecionado.</div>
          </div>
        </div>
      </div>

      <div class="mt-4">
        <h2 class="h6">Prévia</h2>
        <pre id="preview" class="p-3 bg-light border rounded small"></pre>
      </div>
    </div>
  `;

  // ===== DOM =====
  const fileInput = container.querySelector("#fileInput");
  const inputTypeEl = container.querySelector("#inputType");
  const btnPreview = container.querySelector("#btnPreview");
  const btnGenerate = container.querySelector("#btnGenerate");
  const statusEl = container.querySelector("#status");
  const previewEl = container.querySelector("#preview");

  const sessionNumberEl = container.querySelector("#sessionNumber");
  const sessionTypeEl = container.querySelector("#sessionType");
  const sessionDateEl = container.querySelector("#sessionDate");

  // ===== Estado do módulo =====
  let rows = null;
  let inputType = "XLS";

  // ===== Configs =====
  const FONT = "Roboto";
  const SIZE_HEADER = 22; // 11pt
  const SIZE_BODY = 20; // 10pt

  // Espaçamentos (tweak aqui)
  const SPACE_AFTER_PROCESS_LINE = 120;
  const SPACE_AFTER_ORGAO = 80;
  const SPACE_AFTER_TIPO = 80;
  const SPACE_AFTER_INTERESSADO = 60;
  const SPACE_AFTER_ADV = 50;

  // ===== Helpers locais =====
  function setStatus(msg) {
    statusEl.textContent = msg;
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

  function inferSistemaFromProcesso(processoRaw) {
    const processo = String(processoRaw ?? "").trim().replace(/\s+/g, "");
    return /^\d{7}-\d$/.test(processo) ? "AP" : "E-TCE";
  }

  function mapPdfStructuredLinesToRows(lines) {
    const processPattern = /(?:\d{7}\s*-\s*\d|\d{5,}\s*[./-]\s*\d{1,4}|\d{6,})/;
    const toUpper = (v) => upper(v).normalize("NFD").replace(/[\u0300-\u036f]/g, "");

    let processoX = 120;
    let orgaoX = 280;
    let modalidadeX = 430;

    for (const line of lines) {
      for (const cell of line.cells) {
        const t = toUpper(cell.text);
        if (t.includes("PROCESSO")) processoX = cell.x;
        if (t.includes("ORGAO") || t.includes("INTERESSADO")) orgaoX = cell.x;
        if (t.includes("MODALIDADE") || t === "TIPO" || t.includes("EXERCICIO")) {
          modalidadeX = cell.x;
        }
      }
    }

    const bound1 = (processoX + orgaoX) / 2;
    const bound2 = (orgaoX + modalidadeX) / 2;

    const parsed = [];
    let currentRelator = "";
    let currentRow = null;
    let col2Lines = [];
    let col3Lines = [];

    function normalizeProcesso(p) {
      return String(p ?? "").replace(/\s*([./-])\s*/g, "$1").trim();
    }

    function finalizeCurrentRow() {
      if (!currentRow?.Processo) return;

      const col2 = col2Lines.filter(Boolean);
      const col3 = col3Lines.filter(Boolean);
      const merged = [...col2, ...col3].filter(Boolean);
      const adv = col2.filter((v) => /^(ADV|ADVOGADO)/i.test(v));
      const isYear = (v) => /^\d{4}$/.test(String(v).trim());
      const isOrgaoLike = (v) =>
        /(PREFEITURA|CAMARA|TRIBUNAL|FUNDO|SECRETARIA|INSTITUTO|MUNICIP|ESTADO|GOVERNO|CONSORCIO|AUTARQUIA)/i.test(
          v
        );
      const isTipoLike = (v) =>
        /(RECURSO|PRESTACAO|AUDITORIA|TOMADA|DENUNCIA|APOSENTADORIA|ADMISSAO|CONSULTA|EMBARGOS|AGRAVO|REPRESENTACAO|PROCESSO)/i.test(
          v
        );

      const orgaoCandidate =
        merged.find((v) => isOrgaoLike(v)) || col2.find((v) => !isYear(v)) || "";

      const tipoFromLabel = (
        col3.find((v) => /TIPO\s*[:\-]/i.test(v)) || ""
      ).replace(/.*TIPO\s*[:\-]\s*/i, "");
      const tipoCandidate =
        tipoFromLabel ||
        merged.find((v) => isTipoLike(v) && !isYear(v) && v !== orgaoCandidate) ||
        col3.find((v) => !isYear(v) && v !== orgaoCandidate) ||
        "";

      const interessados = col2.filter(
        (v) =>
          !/^(ADV|ADVOGADO)/i.test(v) &&
          v !== orgaoCandidate &&
          v !== tipoCandidate &&
          !isYear(v)
      );

      currentRow["Órgão"] = currentRow["Órgão"] || orgaoCandidate;
      currentRow.Interessados = currentRow.Interessados || interessados.join("\n");
      currentRow["Tipo Processo"] =
        currentRow["Tipo Processo"] ||
        tipoCandidate;
      currentRow.Advogados = currentRow.Advogados || adv.join("\n");

      parsed.push(currentRow);
      currentRow = null;
      col2Lines = [];
      col3Lines = [];
    }

    for (const line of lines) {
      const processTokens = [];
      const col2Tokens = [];
      const col3Tokens = [];

      for (const cell of line.cells) {
        if (cell.x <= bound1) processTokens.push(cell.text);
        else if (cell.x <= bound2) col2Tokens.push(cell.text);
        else col3Tokens.push(cell.text);
      }

      const col1Text = processTokens.join(" ").trim();
      const col2Text = col2Tokens.join(" ").trim();
      const col3Text = col3Tokens.join(" ").trim();
      const normalizedLine = toUpper(line.text);

      const relatorMatch = line.text.match(/RELATOR(?:A)?\s*:\s*(.+)$/i);
      if (relatorMatch) {
        currentRelator = relatorMatch[1].trim();
        continue;
      }
      if (!currentRelator && normalizedLine.startsWith("RELATOR")) {
        continue;
      }

      const processMatch = col1Text.match(processPattern) || line.text.match(processPattern);
      if (processMatch) {
        finalizeCurrentRow();
        currentRow = {
          Relator: currentRelator,
          Processo: normalizeProcesso(processMatch[0]),
          Órgão: "",
          "Tipo Processo": "",
          Interessados: "",
          Advogados: "",
        };
        continue;
      }

      if (!currentRow) continue;

      if (col2Text && !/^(PROCESSO|ORGAO|INTERESSADO)$/i.test(toUpper(col2Text))) {
        col2Lines.push(col2Text);
      }
      if (col3Text && !/^(MODALIDADE|TIPO|EXERCICIO)$/i.test(toUpper(col3Text))) {
        col3Lines.push(col3Text);
      }
    }

    finalizeCurrentRow();
    return parsed;
  }

  function mapPdfTextToRows(rawText) {
    const text = String(rawText ?? "").replace(/\r/g, "");
    const lines = text
      .split("\n")
      .map((l) => l.replace(/\s+/g, " ").trim())
      .filter(Boolean);

    const processPattern = /(?:\d{7}\s*-\s*\d|\d{5,}\s*[./-]\s*\d{1,4}|\d{6,})/;

    const parsed = [];
    let currentRelator = "";
    let currentRow = null;
    let currentField = "";
    let currentMiscLines = [];
    let expectingRelatorName = false;
    let expectingProcessNumber = false;

    function pushCurrentRow() {
      if (!currentRow) return;
      if (!String(currentRow.Processo || "").trim()) return;

      if (!currentRow["Órgão"] && currentMiscLines.length) {
        currentRow["Órgão"] = currentMiscLines.shift();
      }
      if (!currentRow["Tipo Processo"] && currentMiscLines.length) {
        currentRow["Tipo Processo"] = currentMiscLines.shift();
      }
      if (!currentRow.Interessados && currentMiscLines.length) {
        currentRow.Interessados = currentMiscLines.join("\n");
      }

      parsed.push(currentRow);
      currentMiscLines = [];
    }

    function ensureRow() {
      if (currentRow) return;
      currentRow = {
        Relator: currentRelator,
        Processo: "",
        Órgão: "",
        "Tipo Processo": "",
        Interessados: "",
        Advogados: "",
      };
    }

    for (const line of lines) {
      const upLine = upper(line).replace(":", "");

      if (upLine.startsWith("RELATOR")) {
        const sameLine = line.split(":")[1]?.trim() || "";
        currentRelator = sameLine;
        expectingRelatorName = !sameLine;
        currentField = "";
        continue;
      }

      if (expectingRelatorName) {
        currentRelator = line;
        expectingRelatorName = false;
        continue;
      }

      if (upLine === "PROCESSO" || upLine === "Nº" || upLine === "N°") {
        expectingProcessNumber = true;
        currentField = "";
        continue;
      }

      const procMatch = line.match(processPattern);
      if (expectingProcessNumber && procMatch) {
        pushCurrentRow();
        currentRow = {
          Relator: currentRelator,
          Processo: procMatch[0].replace(/\s*([./-])\s*/g, "$1"),
          Órgão: "",
          "Tipo Processo": "",
          Interessados: "",
          Advogados: "",
        };
        expectingProcessNumber = false;
        currentField = "";
        currentMiscLines = [];
        continue;
      }

      if (upLine.startsWith("ÓRGÃO") || upLine.startsWith("ORGAO")) {
        ensureRow();
        currentField = "Órgão";
        const value = line.split(":").slice(1).join(":").trim();
        if (value) currentRow["Órgão"] = value;
        continue;
      }

      if (upLine.startsWith("TIPO PROCESSO")) {
        ensureRow();
        currentField = "Tipo Processo";
        const value = line.split(":").slice(1).join(":").trim();
        if (value) currentRow["Tipo Processo"] = value;
        continue;
      }

      if (upLine.startsWith("INTERESSADO")) {
        ensureRow();
        currentField = "Interessados";
        const value = line.split(":").slice(1).join(":").trim();
        if (value) currentRow.Interessados = value;
        continue;
      }

      if (upLine.startsWith("ADVOGADO")) {
        ensureRow();
        currentField = "Advogados";
        const value = line.split(":").slice(1).join(":").trim();
        if (value) currentRow.Advogados = value;
        continue;
      }

      if (!currentRow && procMatch) {
        pushCurrentRow();
        currentRow = {
          Relator: currentRelator,
          Processo: procMatch[0].replace(/\s*([./-])\s*/g, "$1"),
          Órgão: "",
          "Tipo Processo": "",
          Interessados: "",
          Advogados: "",
        };
        currentField = "";
        currentMiscLines = [];
        continue;
      }

      if (currentRow && currentField) {
        currentRow[currentField] = currentRow[currentField]
          ? `${currentRow[currentField]}\n${line}`
          : line;
      } else if (currentRow) {
        currentMiscLines.push(line);
      }
    }

    pushCurrentRow();

    if (!parsed.length) throw new Error("Não foi possível identificar os processos no PDF.");
    return parsed;
  }

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
    const isConselheiro = CONSELHEIROS.some((key) => n.includes(key));
    return isConselheiro ? "CONSELHEIRO" : "CONSELHEIRO SUBSTITUTO";
  }

  // ===== Listeners (guardados pra destroy) =====
  const listeners = [];

  function on(el, evt, fn) {
    el.addEventListener(evt, fn);
    listeners.push(() => el.removeEventListener(evt, fn));
  }

  updateButtons();

  [sessionNumberEl, sessionTypeEl, sessionDateEl].forEach((el) => {
    on(el, "input", updateButtons);
    on(el, "change", updateButtons);
  });

  on(inputTypeEl, "change", () => {
    inputType = inputTypeEl.value === "PDF" ? "PDF" : "XLS";
    fileInput.value = "";
    rows = null;
    previewEl.textContent = "";
    fileInput.accept = inputType === "PDF" ? ".pdf" : ".xlsx";
    setStatus(
      inputType === "PDF"
        ? "Tipo PDF selecionado. Escolha um arquivo .pdf."
        : "Tipo XLS selecionado. Escolha um arquivo .xlsx."
    );
    updateButtons();
  });

  // ===== Leitura do XLSX =====
  on(fileInput, "change", async (e) => {
    previewEl.textContent = "";
    rows = null;
    updateButtons();

    const file = e.target.files?.[0];
    if (!file) {
      setStatus("Nenhum arquivo selecionado.");
      return;
    }

    setStatus(inputType === "PDF" ? "Lendo PDF..." : "Lendo XLSX...");

    try {
      if (inputType === "PDF") {
        const structured = await readPdfToStructuredLines(file);
        rows = mapPdfStructuredLinesToRows(structured);
        if (!rows.length) {
          rows = mapPdfTextToRows(await readPdfToText(file));
        }
      } else {
        rows = await readFirstSheetXlsxToJson(file);
      }
      if (!rows?.length) {
        throw new Error("Não foi possível extrair processos do arquivo informado.");
      }
      setStatus(`${inputType} OK. Linhas: ${rows.length}.`);
      updateButtons();
    } catch (err) {
      console.error(err);
      setStatus(
        err?.message ||
          `Erro ao ler ${inputType}. Abra o Console (F12) e veja o erro.`
      );
      rows = null;
      updateButtons();
    }
  });

  // ===== Prévia =====
  on(btnPreview, "click", () => {
    if (!rows) return;
    previewEl.textContent = JSON.stringify(rows.slice(0, 10), null, 2);
  });

  // ===== DOCX =====
  on(btnGenerate, "click", async () => {
    if (!rows) return;
    if (!headerOk()) {
      setStatus("Preencha Nº da sessão, Tipo de sessão e Data.");
      return;
    }
    if (!window.docx || !window.saveAs) {
      setStatus("Bibliotecas docx/FileSaver não carregadas (CDN).");
      return;
    }

    setStatus("Gerando DOCX...");

    try {
      const { Document, Packer, Paragraph, TextRun, AlignmentType } = window.docx;

      const sessionNumber = ordinalFeminino(sessionNumberEl.value);
      const sessionType = upper(sessionTypeEl.value);
      const dateBR = formatDateBR(sessionDateEl.value);

      const children = [];

      const separator = () =>
        new Paragraph({
          children: [
            new TextRun(
              "______________________________________________________________________________________"
            ),
          ],
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

        // Linha em branco entre relator e primeiro processo
        children.push(blankLine(120));

        for (const row of processos) {
          const sistema =
            inputType === "PDF"
              ? inferSistemaFromProcesso(row["Processo"])
              : upper(row["Sistema de Tramitação"]) ||
                inferSistemaFromProcesso(row["Processo"]);
          const processo = String(row["Processo"] ?? "").trim();

          let label = "PROCESSO";
          let color = "000000";
          if (sistema === "E-TCE") {
            label = "PROCESSO ELETRÔNICO";
            color = "FF0000";
          } else if (sistema === "AP") {
            label = "PROCESSO DIGITAL";
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
                  text: `Nº ${processo}`,
                  bold: true,
                  color: "000000",
                  size: SIZE_BODY,
                  font: FONT,
                }),
              ],
              spacing: { after: SPACE_AFTER_PROCESS_LINE },
            })
          );

          const orgao = upper(row["Órgão"]);
          const tipoProcesso = upper(row["Tipo Processo"]);

          if (orgao) {
            // Órgão: em negrito
            children.push(
              new Paragraph({
                children: [
                  new TextRun({
                    text: orgao,
                    bold: true,
                    size: SIZE_BODY,
                    font: FONT,
                  }),
                ],
                spacing: { after: SPACE_AFTER_ORGAO },
              })
            );
          }

          if (tipoProcesso) {
            // Tipo Processo: negrito
            children.push(
              new Paragraph({
                children: [
                  new TextRun({
                    text: tipoProcesso,
                    bold: true,
                    size: SIZE_BODY,
                    font: FONT,
                  }),
                ],
                spacing: { after: SPACE_AFTER_TIPO },
              })
            );
          }

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

          // Espaço entre processos
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

  // interface do módulo
  return {
    destroy() {
      // remove listeners
      listeners.forEach((off) => off());
      // limpa o container (opcional)
      // container.innerHTML = "";
    },
  };
}
