/* global XLSX, docx, saveAs */

const fileInput = document.getElementById("fileInput");
const btnPreview = document.getElementById("btnPreview");
const btnGenerate = document.getElementById("btnGenerate");
const statusEl = document.getElementById("status");
const previewEl = document.getElementById("preview");

let workbook = null;
let rows = null; // array de objetos (linha -> colunas)

function setStatus(msg) {
  statusEl.textContent = msg;
}

function clearPreview() {
  previewEl.textContent = "";
}

fileInput.addEventListener("change", async (e) => {
  clearPreview();
  rows = null;
  workbook = null;

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
    workbook = XLSX.read(arrayBuffer, { type: "array" });

    const firstSheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[firstSheetName];

    // Converte para array de objetos usando a primeira linha como cabeçalho
    rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    btnPreview.disabled = false;
    btnGenerate.disabled = false;

    setStatus(
      `OK: ${file.name} | Aba: "${firstSheetName}" | Linhas: ${rows.length}`
    );
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

btnGenerate.addEventListener("click", async () => {
  if (!rows) return;

  setStatus("Gerando DOCX...");

  try {
    // EXEMPLO: cria um documento com um título e uma tabela simples com as colunas do XLSX
    const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell } =
      docx;

    const columns = rows.length > 0 ? Object.keys(rows[0]) : [];

    const title = new Paragraph({
      children: [
        new TextRun({ text: "Relatório gerado", bold: true, size: 28 }),
      ],
      spacing: { after: 300 },
    });

    const headerRow = new TableRow({
      children: columns.map(
        (col) =>
          new TableCell({
            children: [new Paragraph({ children: [new TextRun({ text: col, bold: true })] })],
          })
      ),
    });

    const dataRows = rows.map((r) => {
      return new TableRow({
        children: columns.map((col) => {
          const value = String(r[col] ?? "");
          return new TableCell({
            children: [new Paragraph(value)],
          });
        }),
      });
    });

    const table = new Table({
      rows: [headerRow, ...dataRows],
    });

    const doc = new Document({
      sections: [
        {
          children: [title, table],
        },
      ],
    });

    const blob = await Packer.toBlob(doc);

    const filename = `relatorio_${new Date().toISOString().slice(0, 10)}.docx`;
    saveAs(blob, filename);

    setStatus(`DOCX gerado: ${filename}`);
  } catch (err) {
    console.error(err);
    setStatus("Erro ao gerar DOCX. Veja o console do navegador.");
  }
});
