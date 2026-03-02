/* global docx, saveAs */

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

          <label for="fileInput" class="form-label">Selecione o arquivo .xlsx</label>
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

  // ===== DOM =====
  const fileInput = container.querySelector("#fileInput");
  const btnPreview = container.querySelector("#btnPreview");
  const btnGenerate = container.querySelector("#btnGenerate");
  const statusEl = container.querySelector("#status");
  const previewEl = container.querySelector("#preview");

  const sessionNumberEl = container.querySelector("#sessionNumber");
  const sessionTypeEl = container.querySelector("#sessionType");
  const sessionDateEl = container.querySelector("#sessionDate");

  // ===== Estado do módulo =====
  let rows = null;

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

    setStatus("Lendo XLSX...");

    try {
      rows = await readFirstSheetXlsxToJson(file);
      setStatus(`XLSX OK. Linhas: ${rows.length}.`);
      updateButtons();
    } catch (err) {
      console.error(err);
      setStatus(err?.message || "Erro ao ler XLSX. Abra o Console (F12) e veja o erro.");
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

          // Órgão: em negrito
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
