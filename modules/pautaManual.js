/* global docx, saveAs, pdfjsLib */

import {
  readFirstSheetXlsxToJson,
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

          <div class="row g-3 mb-3">
            <div class="col-md-4">
              <label for="inputMode" class="form-label">Tipo de arquivo de entrada</label>
              <select id="inputMode" class="form-select">
                <option value="XLSX" selected>XLSX</option>
                <option value="PDF">PDF</option>
              </select>
            </div>
          </div>

          <label for="fileInput" class="form-label" id="fileInputLabel">
            Selecione o arquivo .xlsx
          </label>
          <input class="form-control" type="file" id="fileInput" accept=".xlsx" />

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

  const fileInput = container.querySelector("#fileInput");
  const fileInputLabel = container.querySelector("#fileInputLabel");
  const inputModeEl = container.querySelector("#inputMode");

  const btnPreview = container.querySelector("#btnPreview");
  const btnGenerate = container.querySelector("#btnGenerate");
  const statusEl = container.querySelector("#status");
  const previewEl = container.querySelector("#preview");

  const sessionNumberEl = container.querySelector("#sessionNumber");
  const sessionTypeEl = container.querySelector("#sessionType");
  const sessionDateEl = container.querySelector("#sessionDate");

  let rows = null;
  let currentInputMode = "XLSX";

  const FONT = "Roboto";
  const SIZE_HEADER = 22;
  const SIZE_BODY = 20;

  const SPACE_AFTER_PROCESS_LINE = 120;
  const SPACE_AFTER_ORGAO = 80;
  const SPACE_AFTER_TIPO = 80;
  const SPACE_AFTER_INTERESSADO = 60;
  const SPACE_AFTER_ADV = 50;

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

  function resetLoadedData() {
    rows = null;
    previewEl.textContent = "";
    fileInput.value = "";
    updateButtons();
  }

  function updateInputModeUI() {
    currentInputMode = inputModeEl.value || "XLSX";

    if (currentInputMode === "PDF") {
      fileInputLabel.textContent = "Selecione o arquivo .pdf";
      fileInput.accept = ".pdf,application/pdf";
      setStatus("Selecione um PDF para extração.");
    } else {
      fileInputLabel.textContent = "Selecione o arquivo .xlsx";
      fileInput.accept = ".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
      setStatus("Nenhum arquivo selecionado.");
    }

    resetLoadedData();
  }

  function normalizeWhitespace(s) {
    return String(s ?? "")
      .replace(/\u00A0/g, " ")
      .replace(/[ \t]+/g, " ")
      .replace(/\s+$/g, "")
      .trim();
  }

  function toIsoDateFromBR(dateBR) {
    const m = String(dateBR || "").match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if (!m) return "";
    const [, dd, mm, yyyy] = m;
    return `${yyyy}-${mm}-${dd}`;
  }

  function inferSessionType(raw) {
    const s = upper(raw);
    if (s.includes("PRIMEIRA CÂMARA")) return "PRIMEIRA CÂMARA";
    if (s.includes("SEGUNDA CÂMARA")) return "SEGUNDA CÂMARA";
    if (s.includes("PLENO")) return "PLENO";
    return "";
  }

  function normalizeExtractedLines(lines) {
    const out = [];

    for (let raw of lines) {
      let line = String(raw ?? "")
        .replace(/\u00A0/g, " ")
        .replace(/[ \t]+/g, " ")
        .trim();

      if (!line) continue;

      line = line.replace(/\s*(RELATOR:\s*)/gi, "\n$1");
      line = line.replace(/(\d{4})(PROCESSO\b)/gi, "$1\n$2");
      line = line.replace(/(\d{7,8}-\d)(PROCESSO\b)/gi, "$1\n$2");
      line = line.replace(/(\d{4})(RELATOR:\s*)/gi, "$1\n$2");
      line = line.replace(/([A-Za-zÀ-ÿ])(\d{7,8}-\d\b)/g, "$1\n$2");
      line = line.replace(/(\d{7,8}-\d)\s+(\d{7,8}-\d\b)/g, "$1\n$2");

      const parts = line
        .split("\n")
        .map((p) => p.trim())
        .filter(Boolean);

      out.push(...parts);
    }

    return out;
  }

  function isHeaderNoise(line) {
    const s = upper(normalizeWhitespace(line));
    return (
      s === "PROCESSO ÓRGÃO / INTERESSADO MODALIDADE / TIPO /" ||
      s === "PROCESSO ÓRGAO / INTERESSADO MODALIDADE / TIPO /" ||
      s === "MODALIDADE / TIPO / EXERCÍCIO" ||
      s === "MODALIDADE / TIPO / EXERCICIO" ||
      s === "ÓRGÃO / INTERESSADO" ||
      s === "ÓRGAO / INTERESSADO" ||
      s === "EXERCÍCIO" ||
      s === "EXERCICIO"
    );
  }

  function isFooterNoise(line) {
    const s = upper(normalizeWhitespace(line));
    return (
      /^RECIFE,\s*\d{1,2}\s+DE\s+[A-ZÇÃÉÊÍÓÔÚ]+\s+DE\s+\d{4}\.?$/i.test(s) ||
      s === "DIRETORIA DE PLENÁRIO" ||
      s === "DIRETORIA DE PLENARIO"
    );
  }

  function matchProcessStart(line) {
    const s = normalizeWhitespace(line);

    let m = s.match(/^(\d{7,8}-\d)\s*(.*)$/);
    if (m) {
      return {
        processo: m[1],
        rest: normalizeWhitespace(m[2] || ""),
      };
    }

    m = s.match(/(\d{7,8}-\d)\s*(.*)$/);
    if (m) {
      return {
        processo: m[1],
        rest: normalizeWhitespace(m[2] || ""),
      };
    }

    return null;
  }

  function isYearLine(line) {
    return /^(19|20)\d{2}$/.test(normalizeWhitespace(line));
  }

  function isLawyerLine(line) {
    return /^(Adv\.|ADV\.|Procurador Habilitado:)/i.test(normalizeWhitespace(line));
  }

  function looksLikeOrgaoContinuation(currentOrgaoLines, nextLine) {
    const prev = normalizeWhitespace(currentOrgaoLines.join(" "));
    const line = normalizeWhitespace(nextLine);

    if (!line) return false;
    if (isLawyerLine(line)) return false;

    if (/(^| )(de|da|do|das|dos|e|em|para|por|ao|aos|à|às|n[oa]s?)$/i.test(prev)) {
      return true;
    }

    if (prev.length >= 35 && line.length <= 35 && !/[0-9]/.test(line)) {
      const upperLine = upper(line);
      const hasInstitutionWord =
        /(PERNAMBUCO|RECIFE|OLINDA|CARUARU|PETROLINA|ARAÇOIABA|MUNICIPAL|ESTADUAL|SECRETARIA|FUNDO|DEPARTAMENTO|UNIVERSIDADE|TRIBUNAL|CÂMARA|CAMARA|PREFEITURA|FUNDAÇÃO|FUNDACAO|INSTITUTO|NÚCLEO|NUCLEO|POLÍCIA|POLICIA)/i.test(
          upperLine,
        );
      if (hasInstitutionWord) return true;
    }

    return false;
  }

  function joinBrokenOabLines(lines) {
    const result = [];

    for (let i = 0; i < lines.length; i++) {
      const current = normalizeWhitespace(lines[i]);
      const next = normalizeWhitespace(lines[i + 1] || "");

      if (/OAB:\s*$/i.test(current) && next) {
        result.push(`${current} ${next}`);
        i += 1;
        continue;
      }

      result.push(current);
    }

    return result;
  }

  function classifySistemaTramitacaoByProcesso(processo) {
    const p = String(processo || "").trim();
    return /^\d{7}-\d$/.test(p) ? "AP" : "E-TCE";
  }

  function extractSessionInfoFromPdfLines(lines) {
    const joined = lines.slice(0, 12).join(" ");
    const normalized = normalizeWhitespace(joined);

    const dateMatch = normalized.match(/DO DIA\s+(\d{2}\/\d{2}\/\d{4})/i);
    const typeMatch = normalized.match(
      /PAUTA DA SESSÃO ORDINÁRIA DA\s+(PLENO|PRIMEIRA CÂMARA|SEGUNDA CÂMARA)/i,
    );

    return {
      sessionType: inferSessionType(typeMatch?.[1] || ""),
      sessionDateIso: toIsoDateFromBR(dateMatch?.[1] || ""),
    };
  }

  async function readPdfToLines(file) {
    if (!window.pdfjsLib) {
      throw new Error("Biblioteca pdf.js não carregada.");
    }

    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;

    const allLines = [];

    for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
      const page = await pdf.getPage(pageNum);
      const textContent = await page.getTextContent();

      const items = textContent.items || [];
      let currentLine = "";

      for (const item of items) {
        const str = String(item.str ?? "")
          .replace(/\u00A0/g, " ")
          .replace(/[ \t]+/g, " ")
          .trim();

        if (!str) continue;

        if (currentLine) {
          currentLine += " " + str;
        } else {
          currentLine = str;
        }

        if (item.hasEOL) {
          allLines.push(currentLine.trim());
          currentLine = "";
        }
      }

      if (currentLine.trim()) {
        allLines.push(currentLine.trim());
      }
    }

    return normalizeExtractedLines(allLines);
  }

  function buildRowFromPdfBlock({ processo, relator, blockLines }) {
    let lines = blockLines
      .map((l) => normalizeWhitespace(l))
      .filter(Boolean);

    lines = joinBrokenOabLines(lines);
    lines = lines.filter((l) => !isHeaderNoise(l) && !isFooterNoise(l));

    let actualYearIdx = -1;
    for (let idx = lines.length - 1; idx >= 0; idx--) {
      if (/^(19|20)\d{2}$/.test(lines[idx])) {
        actualYearIdx = idx;
        break;
      }
    }

    if (actualYearIdx < 2) {
      throw new Error(`Não foi possível identificar modalidade/tipo/exercício do processo ${processo}.`);
    }

    const exercicio = lines[actualYearIdx];
    const tipo = lines[actualYearIdx - 1];
    const modalidade = lines[actualYearIdx - 2];

    const middle = lines.slice(0, actualYearIdx - 2);

    if (!middle.length) {
      throw new Error(`Não foi possível identificar órgão/interessados do processo ${processo}.`);
    }

    const orgaoLines = [middle[0]];
    const interessadosBrutos = [];

    for (let i = 1; i < middle.length; i++) {
      const line = middle[i];

      if (!interessadosBrutos.length && looksLikeOrgaoContinuation(orgaoLines, line)) {
        orgaoLines.push(line);
      } else {
        interessadosBrutos.push(line);
      }
    }

    const advogados = [];
    const interessados = [];

    for (const item of interessadosBrutos) {
      if (isLawyerLine(item)) {
        advogados.push(
          item
            .replace(/^Adv\.\s*/i, "")
            .replace(/^Procurador Habilitado:\s*/i, "Procurador Habilitado: ")
            .trim(),
        );
      } else {
        interessados.push(item);
      }
    }

    return {
      Processo: processo,
      Relator: relator,
      Órgão: orgaoLines.join(" "),
      "Tipo Processo": tipo,
      Modalidade: modalidade,
      Exercício: exercicio,
      Interessados: interessados.join("\n"),
      Advogados: advogados.join("\n"),
      "Sistema de Tramitação": classifySistemaTramitacaoByProcesso(processo),
      Voto: "",
    };
  }

  function parseRowsFromPdfLines(lines) {
    const rowsOut = [];
    let currentRelator = "";
    let i = 0;

    while (i < lines.length) {
      const rawLine = lines[i];
      const line = normalizeWhitespace(rawLine);

      if (!line || isHeaderNoise(line) || isFooterNoise(line)) {
        i++;
        continue;
      }

      if (/^RELATOR:/i.test(line)) {
        currentRelator = normalizeWhitespace(
          line.replace(/^RELATOR:\s*/i, "").replace(/\s+/g, " "),
        );
        i++;
        continue;
      }

      const processStart = matchProcessStart(line);

      if (processStart) {
        const blockLines = [];
        if (processStart.rest) blockLines.push(processStart.rest);

        i++;

        while (i < lines.length) {
          const nextLine = normalizeWhitespace(lines[i]);

          if (!nextLine) {
            i++;
            continue;
          }

          if (isHeaderNoise(nextLine) || isFooterNoise(nextLine)) {
            i++;
            continue;
          }

          if (/^RELATOR:/i.test(nextLine)) {
            break;
          }

          if (matchProcessStart(nextLine)) {
            break;
          }

          blockLines.push(nextLine);

          if (/\b(19|20)\d{2}\b/.test(nextLine)) {
            i++;
            break;
          }

          i++;
        }

        try {
          const row = buildRowFromPdfBlock({
            processo: processStart.processo,
            relator: currentRelator,
            blockLines,
          });
          rowsOut.push(row);
        } catch (err) {
          console.error("Falha ao montar bloco do processo:", processStart.processo, blockLines, err);
        }

        continue;
      }

      i++;
    }

    return rowsOut;
  }

  async function readPdfToRows(file) {
    const lines = await readPdfToLines(file);
    const sessionInfo = extractSessionInfoFromPdfLines(lines);
    const rowsParsed = parseRowsFromPdfLines(lines);

    console.log("PDF lines:", lines);
    console.log("PDF rows parsed:", rowsParsed);

    return {
      rows: rowsParsed,
      sessionInfo,
      lines,
    };
  }

  function fillSessionFieldsIfEmpty(sessionInfo) {
    if (!sessionTypeEl.value && sessionInfo.sessionType) {
      sessionTypeEl.value = sessionInfo.sessionType;
    }

    if (!sessionDateEl.value && sessionInfo.sessionDateIso) {
      sessionDateEl.value = sessionInfo.sessionDateIso;
    }
  }

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

  const listeners = [];

  function on(el, evt, fn) {
    el.addEventListener(evt, fn);
    listeners.push(() => el.removeEventListener(evt, fn));
  }

  updateButtons();
  updateInputModeUI();

  [sessionNumberEl, sessionTypeEl, sessionDateEl].forEach((el) => {
    on(el, "input", updateButtons);
    on(el, "change", updateButtons);
  });

  on(inputModeEl, "change", () => {
    updateInputModeUI();
  });

  on(fileInput, "change", async (e) => {
    previewEl.textContent = "";
    rows = null;
    updateButtons();

    const file = e.target.files?.[0];
    if (!file) {
      setStatus("Nenhum arquivo selecionado.");
      return;
    }

    try {
      if (currentInputMode === "PDF") {
        setStatus("Lendo PDF...");

        const result = await readPdfToRows(file);
        rows = result.rows || [];
        fillSessionFieldsIfEmpty(result.sessionInfo);

        if (!rows.length) {
          setStatus("PDF lido, mas nenhum processo foi identificado.");
          rows = null;
          updateButtons();
          return;
        }

        setStatus(`PDF OK. Processos identificados: ${rows.length}.`);
        updateButtons();
        return;
      }

      setStatus("Lendo XLSX...");
      rows = await readFirstSheetXlsxToJson(file);

      setStatus(`XLSX OK. Linhas: ${rows.length}.`);
      updateButtons();
    } catch (err) {
      console.error(err);
      setStatus(err?.message || `Erro ao ler ${currentInputMode}. Abra o Console (F12) e veja o erro.`);
      rows = null;
      updateButtons();
    }
  });

  on(btnPreview, "click", () => {
    if (!rows) return;
    previewEl.textContent = JSON.stringify(rows.slice(0, 10), null, 2);
  });

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

      const grouped = new Map();
      for (const r of rows) {
        const rel = String(r["Relator"] ?? "").trim();
        if (!grouped.has(rel)) grouped.set(rel, []);
        grouped.get(rel).push(r);
      }

      for (const [relator, processos] of grouped.entries()) {
        const prefix = relatorPrefix(relator);

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

  return {
    destroy() {
      listeners.forEach((off) => off());
    },
  };
}
