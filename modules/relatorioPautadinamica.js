/* global docx, saveAs */

import { readFirstSheetXlsxToJson, formatDateBR } from "../shared/helpers.js";

export function mount(container) {
  container.innerHTML = `
    <div class="module-card">
      <div class="card">
        <div class="card-body">
          <div class="row g-3 mb-3">
            <div class="col-md-4">
              <label for="reportDate" class="form-label">Data de referência</label>
              <input id="reportDate" type="date" class="form-control" />
            </div>
          </div>

          <label for="fileInputR" class="form-label">Selecione o arquivo .xlsx</label>
          <input class="form-control" type="file" id="fileInputR" accept=".xlsx" />

          <div class="d-flex gap-2 mt-3">
            <button id="btnPreviewR" class="btn btn-outline-secondary" disabled>
              Pré-visualizar (10 linhas)
            </button>

            <button id="btnGenerateR" class="btn btn-primary" disabled>
              Gerar DOCX
            </button>
          </div>

          <div class="mt-3">
            <div id="statusR" class="small text-muted">Nenhum arquivo selecionado.</div>
          </div>
        </div>
      </div>

      <div class="mt-4">
        <h2 class="h6">Prévia</h2>
        <pre id="previewR" class="p-3 bg-light border rounded small"></pre>
      </div>
    </div>
  `;

  const fileInput = container.querySelector("#fileInputR");
  const btnPreview = container.querySelector("#btnPreviewR");
  const btnGenerate = container.querySelector("#btnGenerateR");
  const statusEl = container.querySelector("#statusR");
  const previewEl = container.querySelector("#previewR");
  const reportDateEl = container.querySelector("#reportDate");

  let rows = null;

  const listeners = [];
  function on(el, evt, fn) {
    el.addEventListener(evt, fn);
    listeners.push(() => el.removeEventListener(evt, fn));
  }

  function setStatus(msg) {
    statusEl.textContent = msg;
  }

  function updateButtons() {
    btnPreview.disabled = !rows;
    btnGenerate.disabled = !rows;
  }
  updateButtons();

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

  on(btnPreview, "click", () => {
    if (!rows) return;
    previewEl.textContent = JSON.stringify(rows.slice(0, 10), null, 2);
  });

  on(btnGenerate, "click", async () => {
    if (!rows) return;
    if (!window.docx || !window.saveAs) {
      setStatus("Bibliotecas docx/FileSaver não carregadas (CDN).");
      return;
    }

    setStatus("Gerando DOCX...");

    try {
      const { Document, Packer, Paragraph, TextRun, AlignmentType } = window.docx;

      const dateBR = formatDateBR(reportDateEl.value);
      const title = "RELATÓRIO PAUTA DINÂMICA";

      const children = [];

      children.push(
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: title, bold: true, size: 24, font: "Roboto" }),
          ],
          spacing: { after: 200 },
        })
      );

      if (dateBR) {
        children.push(
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({ text: `DATA: ${dateBR}`, bold: true, size: 22, font: "Roboto" }),
            ],
            spacing: { after: 200 },
          })
        );
      }

      children.push(
        new Paragraph({
          children: [
            new TextRun({ text: `Total de linhas no XLSX: ${rows.length}`, size: 20, font: "Roboto" }),
          ],
          spacing: { after: 200 },
        })
      );

      // Dump controlado (pra não explodir o DOCX)
      const limit = Math.min(50, rows.length);
      for (let i = 0; i < limit; i++) {
        children.push(
          new Paragraph({
            children: [
              new TextRun({ text: `${i + 1}. `, bold: true, size: 18, font: "Roboto" }),
              new TextRun({ text: JSON.stringify(rows[i]), size: 18, font: "Roboto" }),
            ],
            spacing: { after: 80 },
          })
        );
      }

      if (rows.length > limit) {
        children.push(
          new Paragraph({
            children: [
              new TextRun({ text: `(...) Exibindo apenas ${limit} de ${rows.length} linhas.`, italics: true, size: 18, font: "Roboto" }),
            ],
            spacing: { before: 120 },
          })
        );
      }

      const doc = new Document({ sections: [{ children }] });
      const blob = await Packer.toBlob(doc);

      const filename = `relatorio_pauta_dinamica${dateBR ? "_" + dateBR.replaceAll("/", "-") : ""}.docx`;
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
