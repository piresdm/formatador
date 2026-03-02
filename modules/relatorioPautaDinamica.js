/* global XLSX, docx, saveAs, pdfMake */

import { splitLines, upper, formatDateBR, ordinalFeminino } from "../shared/helpers.js";

export function mount(container) {
  container.innerHTML = `
    <div class="card">
      <div class="card-body">

        <div class="row g-3">
          <div class="col-md-3">
            <label for="sessionNumberR" class="form-label">Nº da sessão</label>
            <input id="sessionNumberR" type="number" min="1" step="1" class="form-control" placeholder="Ex: 20" required />
          </div>

          <div class="col-md-5">
            <label for="sessionTypeR" class="form-label">Tipo de sessão</label>
            <select id="sessionTypeR" class="form-select" required>
              <option value="" selected>Selecione...</option>
              <option value="PLENO">Pleno</option>
              <option value="1CAM">1ª Câmara</option>
              <option value="2CAM">2ª Câmara</option>
            </select>
          </div>

          <div class="col-md-4">
            <label for="sessionDateR" class="form-label">Data</label>
            <input id="sessionDateR" type="date" class="form-control" required />
          </div>
        </div>

        <hr class="my-3" />

        <div class="row g-3">
          <div class="col-12">
            <label for="fileInputR" class="form-label">Selecione o arquivo .xlsx</label>

            <!-- Wrapper para colocar o "X" dentro do input -->
            <div class="position-relative">
              <input class="form-control pe-5" type="file" id="fileInputR" accept=".xlsx" />
              <button
                id="btnClearFileX"
                type="button"
                class="btn-close position-absolute top-50 end-0 translate-middle-y me-2 d-none"
                aria-label="Remover arquivo"
                title="Remover arquivo"
              ></button>
            </div>

            <div class="form-text" id="fileHintR">Nenhum arquivo selecionado.</div>
          </div>
        </div>

        <div class="d-flex flex-wrap gap-2 mt-3 align-items-center">
          <button id="btnClearAllR" class="btn btn-outline-secondary" type="button">Limpar</button>

          <button id="btnPdf" class="btn btn-primary" disabled>Gerar PDF</button>
          <button id="btnDocx" class="btn btn-outline-primary" disabled>Gerar DOCX</button>
        </div>

      </div>
    </div>
  `;

  // ===== DOM =====
  const fileInput = container.querySelector("#fileInputR");
  const fileHintEl = container.querySelector("#fileHintR");
  const btnClearFileX = container.querySelector("#btnClearFileX");

  const btnPdf = container.querySelector("#btnPdf");
  const btnDocx = container.querySelector("#btnDocx");
  const btnClearAll = container.querySelector("#btnClearAllR");

  const sessionNumberEl = container.querySelector("#sessionNumberR");
  const sessionTypeEl = container.querySelector("#sessionTypeR");
  const sessionDateEl = container.querySelector("#sessionDateR");

  // ===== Estado =====
  let allRows = null;
  let lastFilenameBase = null;
  let logoDataUrl = null;

  // ===== Validação =====
  const STATUS_INICIAIS = new Set([
    "Preferência",
    "Preferência e Sustentação Oral",
    "Destaque Sessão Virtual",
    "Devolução Vista",
    "Retorno de Adiantamento",
    "Pauta",
    "Extrapauta",
  ]);

  const STATUS_FINAIS = new Set([
    "Pedido de Vista",
    "Adiado",
    "Sobrestado",
    "Retirado de Pauta",
    "Julgado",
  ]);

  // ===== Listeners cleanup =====
  const listeners = [];
  function on(el, evt, fn) {
    el.addEventListener(evt, fn);
    listeners.push(() => el.removeEventListener(evt, fn));
  }

  // ===== UX helpers =====
  function userAlert(msg) {
    // Popup simples com OK (sem cara de programador)
    window.alert(String(msg || ""));
  }

  function headerOk() {
    return (
      !!ordinalFeminino(sessionNumberEl.value) &&
      !!String(sessionTypeEl.value || "").trim() &&
      !!String(sessionDateEl.value || "").trim()
    );
  }

  function normSessionType(v) {
    return String(v ?? "").trim().toUpperCase();
  }

  function setFileHint(text) {
    fileHintEl.textContent = text;
  }

  function setClearXVisible(visible) {
    btnClearFileX.classList.toggle("d-none", !visible);
  }

  function updateButtons() {
    const ok = !!allRows && headerOk();
    btnPdf.disabled = !ok;
    btnDocx.disabled = !ok;
    setClearXVisible(!!allRows || !!fileInput.files?.[0]);
  }

  [sessionNumberEl, sessionTypeEl, sessionDateEl].forEach((el) => {
    on(el, "input", updateButtons);
    on(el, "change", updateButtons);
  });

  // ===== Carregar logo do repo (1x) =====
  async function loadLogoOnce() {
    if (logoDataUrl) return logoDataUrl;

    const url = new URL("./assets/logo.png", window.location.href).toString();
    const resp = await fetch(url, { cache: "no-store" });

    if (!resp.ok) {
      // Mensagem para usuário (sem detalhes técnicos)
      throw new Error("Não foi possível carregar a imagem do cabeçalho (logo).");
    }

    const blob = await resp.blob();
    logoDataUrl = await blobToDataURL(blob);
    return logoDataUrl;
  }

  function blobToDataURL(blob) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = reject;
      reader.readAsDataURL(blob);
    });
  }

  // ===== Reset =====
  function clearLoadedFileState({ silent = false } = {}) {
    allRows = null;
    lastFilenameBase = null;

    // limpar o input file
    fileInput.value = "";
    setFileHint("Nenhum arquivo selecionado.");
    setClearXVisible(false);
    updateButtons();

    if (!silent) userAlert("Arquivo removido.");
  }

  function clearAllFields() {
    sessionNumberEl.value = "";
    sessionTypeEl.value = "";
    sessionDateEl.value = "";
    clearLoadedFileState({ silent: true });
    updateButtons();
    userAlert("Campos limpos.");
  }

  // X dentro do input
  on(btnClearFileX, "click", () => clearLoadedFileState());

  // Limpar tudo (discreto ao lado de Gerar)
  on(btnClearAll, "click", clearAllFields);

  // ===== Leitura do XLSX =====
  on(fileInput, "change", async (e) => {
    allRows = null;
    lastFilenameBase = null;
    updateButtons();

    const file = e.target.files?.[0];

    if (!file) {
      setFileHint("Nenhum arquivo selecionado.");
      setClearXVisible(false);
      updateButtons();
      return;
    }

    if (!file.name.toLowerCase().endsWith(".xlsx")) {
      setFileHint("Selecione um arquivo .xlsx.");
      userAlert("Selecione um arquivo .xlsx.");
      clearLoadedFileState({ silent: true });
      return;
    }

    if (!window.XLSX) {
      userAlert("Não foi possível ler a planilha. Tente recarregar a página.");
      clearLoadedFileState({ silent: true });
      return;
    }

    setFileHint(`Arquivo selecionado: ${file.name}`);
    setClearXVisible(true);

    try {
      const arrayBuffer = await file.arrayBuffer();

      // Evita conversão automática para Date (reduz bug de fuso/UTC)
      const workbook = window.XLSX.read(arrayBuffer, { type: "array", cellDates: false });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];

      // raw:true preserva datas numéricas do Excel
      allRows = window.XLSX.utils.sheet_to_json(sheet, { defval: "", raw: true });

      userAlert(`Planilha carregada com sucesso. Linhas: ${allRows.length}.`);
      updateButtons();
    } catch (err) {
      console.error(err);
      userAlert("Não foi possível ler a planilha. Verifique se o arquivo está correto e tente novamente.");
      clearLoadedFileState({ silent: true });
    }
  });

  // ===== Sessão =====
  function mapSessionTypeToHeader(typeCode) {
    if (typeCode === "PLENO") return "PLENO";
    if (typeCode === "1CAM") return "PRIMEIRA CÂMARA";
    if (typeCode === "2CAM") return "SEGUNDA CÂMARA";
    return String(typeCode || "").toUpperCase();
  }

  function excelCellToYmd(value) {
    if (!value) return "";

    // Número do Excel (ideal)
    if (typeof value === "number" && window.XLSX?.SSF?.parse_date_code) {
      const parsed = window.XLSX.SSF.parse_date_code(value);
      if (!parsed) return "";
      const y = parsed.y;
      const m = String(parsed.m).padStart(2, "0");
      const d = String(parsed.d).padStart(2, "0");
      return `${y}-${m}-${d}`;
    }

    // Se vier como Date, usa UTC para não “voltar um dia”
    if (value instanceof Date && !isNaN(value.getTime())) {
      const y = value.getUTCFullYear();
      const m = String(value.getUTCMonth() + 1).padStart(2, "0");
      const d = String(value.getUTCDate()).padStart(2, "0");
      return `${y}-${m}-${d}`;
    }

    const s = String(value).trim();
    const m1 = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if (m1) return `${m1[3]}-${m1[2]}-${m1[1]}`;

    const m2 = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (m2) return s;

    return "";
  }

  function relator2Nomes(nome) {
    const parts = String(nome ?? "").trim().split(/\s+/).filter(Boolean);
    return parts.slice(0, 2).join(" ");
  }

  function buildValidationAlertText(groups, maxItemsTotal = 20) {
    const lines = [];
    let shown = 0;
    let total = 0;

    for (const g of groups) total += g.items.length;

    lines.push(`Foram encontrados erros na planilha (${total}).`);
    lines.push("Ajuste e gere novamente.");

    for (const g of groups) {
      if (!g.items.length) continue;
      lines.push("");
      lines.push(`${g.title}:`);

      for (const it of g.items) {
        if (shown >= maxItemsTotal) break;
        const det = it.detalhe ? ` (${it.detalhe})` : "";
        lines.push(`- ${it.processo} — ${it.relator2} — ${it.orgao}${det}`);
        shown++;
      }
      if (shown >= maxItemsTotal) break;
    }

    if (total > shown) {
      lines.push("");
      lines.push(`(Mostrando ${shown} de ${total} erros)`);
    }

    return lines.join("\n");
  }

  function validateAndGetSessionRows() {
    if (!allRows) {
      return { ok: false, rows: [], message: "Carregue a planilha primeiro." };
    }
    if (!headerOk()) {
      return { ok: false, rows: [], message: "Preencha Nº da sessão, Tipo de sessão e Data." };
    }

    const typeCode = normSessionType(sessionTypeEl.value);
    const ymd = sessionDateEl.value;

    const filtered = allRows.filter((r) => {
      const tipo = normSessionType(r["Tipo Sessão"]);
      const dataYmd = excelCellToYmd(r["Data"]);
      return tipo === typeCode && dataYmd === ymd;
    });

    if (filtered.length === 0) {
      const dateBR = formatDateBR(ymd);
      return {
        ok: false,
        rows: [],
        message:
          `Não foram encontradas sessões para o dia ${dateBR}.\n` +
          `Confira se a coluna "Tipo Sessão" está como ${typeCode} e se a coluna "Data" é a data da sessão.`,
      };
    }

    const groups = [
      { title: "Status Final vazio", items: [] },
      { title: "Status Final inválido", items: [] },
      { title: "Status Inicial vazio", items: [] },
      { title: "Status Inicial inválido", items: [] },
    ];

    for (const r of filtered) {
      const processo = String(r["Processo"] ?? "").trim() || "(sem processo)";
      const relator = String(r["Relator"] ?? "").trim();
      const orgao = String(r["Órgão"] ?? "").trim() || "(sem órgão)";
      const rel2 = relator2Nomes(relator) || "(sem relator)";

      const stIni = String(r["Status"] ?? "").trim();
      const stFim = String(r["Status Final"] ?? "").trim();

      if (!stFim) {
        groups[0].items.push({ processo, relator2: rel2, orgao });
      } else if (!STATUS_FINAIS.has(stFim)) {
        groups[1].items.push({ processo, relator2: rel2, orgao, detalhe: `Valor: "${stFim}"` });
      }

      if (!stIni) {
        groups[2].items.push({ processo, relator2: rel2, orgao });
      } else if (!STATUS_INICIAIS.has(stIni)) {
        groups[3].items.push({ processo, relator2: rel2, orgao, detalhe: `Valor: "${stIni}"` });
      }
    }

    const hasErrors = groups.some((g) => g.items.length > 0);
    if (hasErrors) {
      return {
        ok: false,
        rows: [],
        message: buildValidationAlertText(groups),
      };
    }

    const dateBR = formatDateBR(ymd);
    lastFilenameBase = `relatorio_pauta_${typeCode}_${dateBR.replaceAll("/", "-")}`;
    return { ok: true, rows: filtered, message: "" };
  }

  // ===== Cores =====
  function statusColorHex(statusFinal) {
    if (statusFinal === "Julgado") return "#1B5E20";
    if (statusFinal === "Adiado" || statusFinal === "Pedido de Vista") return "#8A6D00";
    if (statusFinal === "Sobrestado") return "#4A148C";
    if (statusFinal === "Retirado de Pauta") return "#7F1D1D";
    return "#111827";
  }

  // ===== Agrupamento por relator =====
  function groupByRelatorPreserveOrder(rows) {
    const map = new Map();
    const order = [];
    for (const r of rows) {
      const rel = String(r["Relator"] ?? "").trim();
      if (!map.has(rel)) {
        map.set(rel, []);
        order.push(rel);
      }
      map.get(rel).push(r);
    }
    return { map, order };
  }

  // ===== Resumos =====
  function countByField(rows, fieldName) {
    const counts = new Map();
    for (const r of rows) {
      const v = String(r[fieldName] ?? "").trim();
      if (!v) continue;
      counts.set(v, (counts.get(v) || 0) + 1);
    }
    return Array.from(counts.entries())
      .filter(([, n]) => n > 0)
      .sort((a, b) => a[0].localeCompare(b[0], "pt-BR"));
  }

  // ===== PDF =====
  async function generatePdf(rows) {
    if (!window.pdfMake) throw new Error("PDF indisponível no momento.");

    const logo = await loadLogoOnce();

    const sessionNumber = ordinalFeminino(sessionNumberEl.value);
    const sessionTypeHeader = mapSessionTypeToHeader(normSessionType(sessionTypeEl.value));
    const dateBR = formatDateBR(sessionDateEl.value);

    const { map, order } = groupByRelatorPreserveOrder(rows);

    const pre = countByField(rows, "Status");
    const post = countByField(rows, "Status Final");

    const relatorDest = new Map();
    const processoDest = new Map();
    for (const rel of order) {
      relatorDest.set(rel, `REL_${hashId(rel)}`);
      const ps = map.get(rel) || [];
      for (const r of ps) {
        const proc = String(r["Processo"] ?? "").trim();
        if (!proc) continue;
        processoDest.set(proc, `P_${hashId(rel + "|" + proc)}`);
      }
    }

    const content = [];

    content.push({ image: "logo", width: 70, alignment: "left", margin: [0, 0, 0, 8] });

    content.push({
      text: `PAUTA DA ${sessionNumber} SESSÃO ORDINÁRIA DO ${sessionTypeHeader}`,
      alignment: "center",
      bold: true,
      fontSize: 14,
      margin: [0, 0, 0, 6],
    });
    content.push({
      text: `DATA: ${dateBR}`,
      alignment: "center",
      bold: true,
      fontSize: 14,
      margin: [0, 0, 0, 6],
    });
    content.push({
      text: `HORÁRIO: 10h`,
      alignment: "center",
      bold: true,
      fontSize: 14,
      margin: [0, 0, 0, 14],
    });

    content.push({
      columns: [summaryBox("Resumo pré-sessão", pre), summaryBox("Resumo pós-sessão", post)],
      columnGap: 12,
      margin: [0, 0, 0, 14],
    });

    content.push({ text: "ÍNDICE (por Relator)", bold: true, fontSize: 12, margin: [0, 0, 0, 8] });

    for (const rel of order) {
      const relId = relatorDest.get(rel);

      content.push({
        text: `RELATOR: ${upper(rel)}`,
        bold: true,
        margin: [0, 6, 0, 4],
        linkToDestination: relId,
        color: "#111827",
      });

      const ps = map.get(rel) || [];
      for (const r of ps) {
        const proc = String(r["Processo"] ?? "").trim();
        const orgao = String(r["Órgão"] ?? "").trim();
        const stFim = String(r["Status Final"] ?? "").trim();
        const dest = processoDest.get(proc);

        content.push({
          text: `${proc} — ${stFim} — ${orgao}`,
          margin: [12, 0, 0, 2],
          linkToDestination: dest,
          color: "#0F172A",
        });

        content.push({
          text: rel,
          margin: [12, 0, 0, 6],
          fontSize: 9,
          color: "#374151",
        });
      }
    }

    for (const rel of order) {
      const ps = map.get(rel) || [];
      const relId = relatorDest.get(rel);

      content.push({ text: "", pageBreak: "before" });
      content.push({ text: "", id: relId });

      content.push({ text: upper(rel), bold: true, fontSize: 13, margin: [0, 0, 0, 10] });

      for (const r of ps) {
        const proc = String(r["Processo"] ?? "").trim();
        const orgao = String(r["Órgão"] ?? "").trim();
        const stIni = String(r["Status"] ?? "").trim();
        const stFim = String(r["Status Final"] ?? "").trim();
        const modalidade = String(r["Modalidade"] ?? "").trim();
        const tipoProc = String(r["Tipo Processo"] ?? "").trim();
        const sist = String(r["Sistema de Tramitação"] ?? "").trim();
        const voto = String(r["Voto"] ?? "").trim();
        const interessados = splitLines(r["Interessados"]);
        const advogados = splitLines(r["Advogados"]);

        const procId = processoDest.get(proc);

        content.push({ text: "", id: procId });

        content.push({
          stack: [
            { text: proc, bold: true, fontSize: 12, margin: [0, 0, 0, 2] },
            { text: stFim, bold: true, fontSize: 12, color: statusColorHex(stFim), margin: [0, 0, 0, 8] },

            sectionTitle("INFORMAÇÕES GERAIS"),
            keyValue("Relator", rel),
            keyValue("Órgão", orgao),

            sectionTitle("CLASSIFICAÇÃO"),
            keyValue("Modalidade – Tipo Processo", `${modalidade}${modalidade && tipoProc ? " - " : ""}${tipoProc}` || "—"),

            sectionTitle("PARTES"),
            listField("Interessados", interessados),
            listField("Advogados", advogados),

            sectionTitle("TRAMITAÇÃO"),
            keyValue("Sistema de Tramitação", sist || "—"),
            keyValue("Status Inicial → Status Final", `${stIni} → ${stFim}`),
            keyValue("Voto", voto || "—"),
          ],
          margin: [0, 0, 0, 12],
          border: [true, true, true, true],
        });
      }
    }

    const docDefinition = {
      pageSize: "A4",
      pageMargins: [40, 35, 40, 45],
      footer: (currentPage, pageCount) => ({
        text: `Página ${currentPage} de ${pageCount}`,
        alignment: "right",
        margin: [0, 0, 40, 10],
        fontSize: 9,
        color: "#374151",
      }),
      images: { logo },
      defaultStyle: { fontSize: 10, color: "#111827" },
      content,
    };

    const filename = `${lastFilenameBase || "relatorio_pauta"}.pdf`;
    window.pdfMake.createPdf(docDefinition).download(filename);
    return filename;
  }

  function summaryBox(title, entries) {
    const lines = entries.length ? entries.map(([k, n]) => `${k}: ${n}`).join("\n") : "—";
    return {
      width: "*",
      stack: [{ text: title, bold: true, margin: [0, 0, 0, 6] }, { text: lines, fontSize: 10, lineHeight: 1.2 }],
      border: [true, true, true, true],
      margin: [0, 0, 0, 0],
      padding: 8,
    };
  }

  function sectionTitle(t) {
    return { text: t, bold: true, fontSize: 9, color: "#334155", margin: [0, 6, 0, 3] };
  }

  function keyValue(k, v) {
    return { text: [{ text: `${k}: `, bold: true }, { text: String(v ?? "") || "—" }], margin: [0, 0, 0, 2] };
  }

  function listField(label, arr) {
    if (!arr || arr.length === 0) return keyValue(label, "—");
    return {
      stack: [{ text: `${label}:`, bold: true, margin: [0, 0, 0, 2] }, ...arr.map((x) => ({ text: x, margin: [10, 0, 0, 1] }))],
      margin: [0, 0, 0, 2],
    };
  }

  function hashId(s) {
    const str = String(s ?? "");
    let h = 2166136261;
    for (let i = 0; i < str.length; i++) {
      h ^= str.charCodeAt(i);
      h = Math.imul(h, 16777619);
    }
    return (h >>> 0).toString(36);
  }

  // ===== DOCX =====
  async function generateDocx(rows) {
    if (!window.docx || !window.saveAs) throw new Error("DOCX indisponível no momento.");

    const logo = await loadLogoOnce();
    const logoBytes = dataUrlToUint8Array(logo);

    const { Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel, Footer, PageNumber, ImageRun } = window.docx;

    const sessionNumber = ordinalFeminino(sessionNumberEl.value);
    const sessionTypeHeader = mapSessionTypeToHeader(normSessionType(sessionTypeEl.value));
    const dateBR = formatDateBR(sessionDateEl.value);

    const { map, order } = groupByRelatorPreserveOrder(rows);
    const pre = countByField(rows, "Status");
    const post = countByField(rows, "Status Final");

    const children = [];

    children.push(
      new Paragraph({
        children: [new ImageRun({ data: logoBytes, transformation: { width: 110, height: 45 } })],
        alignment: AlignmentType.LEFT,
        spacing: { after: 120 },
      })
    );

    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: `PAUTA DA ${sessionNumber} SESSÃO ORDINÁRIA DO ${sessionTypeHeader}`, bold: true, size: 24 })],
        spacing: { after: 120 },
      })
    );

    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: `DATA: ${dateBR}`, bold: true, size: 24 })],
        spacing: { after: 80 },
      })
    );

    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: `HORÁRIO: 10h`, bold: true, size: 24 })],
        spacing: { after: 180 },
      })
    );

    children.push(new Paragraph({ children: [new TextRun({ text: "RESUMO PRÉ-SESSÃO", bold: true })], spacing: { after: 60 } }));
    children.push(...summaryParagraphsDocx(pre));

    children.push(new Paragraph({ children: [new TextRun({ text: "RESUMO PÓS-SESSÃO", bold: true })], spacing: { before: 120, after: 60 } }));
    children.push(...summaryParagraphsDocx(post));

    children.push(new Paragraph({ children: [new TextRun({ text: "ÍNDICE (por Relator)", bold: true })], spacing: { before: 180, after: 80 } }));

    for (const rel of order) {
      children.push(new Paragraph({ children: [new TextRun({ text: `RELATOR: ${upper(rel)}`, bold: true })], spacing: { before: 80, after: 40 } }));

      const ps = map.get(rel) || [];
      for (const r of ps) {
        const proc = String(r["Processo"] ?? "").trim();
        const orgao = String(r["Órgão"] ?? "").trim();
        const stFim = String(r["Status Final"] ?? "").trim();
        children.push(new Paragraph({ children: [new TextRun({ text: `• ${proc} — ${stFim} — ${orgao}` })], spacing: { after: 20 } }));
      }
    }

    for (const rel of order) {
      const ps = map.get(rel) || [];

      children.push(new Paragraph({ children: [new TextRun("")], pageBreakBefore: true }));
      children.push(new Paragraph({ text: upper(rel), heading: HeadingLevel.HEADING_1, spacing: { after: 160 } }));

      for (const r of ps) {
        const proc = String(r["Processo"] ?? "").trim();
        const orgao = String(r["Órgão"] ?? "").trim();
        const stIni = String(r["Status"] ?? "").trim();
        const stFim = String(r["Status Final"] ?? "").trim();
        const modalidade = String(r["Modalidade"] ?? "").trim();
        const tipoProc = String(r["Tipo Processo"] ?? "").trim();
        const sist = String(r["Sistema de Tramitação"] ?? "").trim();
        const voto = String(r["Voto"] ?? "").trim();
        const interessados = splitLines(r["Interessados"]);
        const advogados = splitLines(r["Advogados"]);

        children.push(new Paragraph({ children: [new TextRun({ text: proc, bold: true, size: 24 })], spacing: { after: 20 } }));
        children.push(new Paragraph({ children: [new TextRun({ text: stFim, bold: true, size: 24, color: hexToDocxColor(statusColorHex(stFim)) })], spacing: { after: 100 } }));

        children.push(sectionDocx("INFORMAÇÕES GERAIS"));
        children.push(kvDocx("Relator", rel));
        children.push(kvDocx("Órgão", orgao || "—"));

        children.push(sectionDocx("CLASSIFICAÇÃO"));
        children.push(kvDocx("Modalidade – Tipo Processo", `${modalidade}${modalidade && tipoProc ? " - " : ""}${tipoProc}` || "—"));

        children.push(sectionDocx("PARTES"));
        children.push(listDocx("Interessados", interessados));
        children.push(listDocx("Advogados", advogados));

        children.push(sectionDocx("TRAMITAÇÃO"));
        children.push(kvDocx("Sistema de Tramitação", sist || "—"));
        children.push(kvDocx("Status Inicial → Status Final", `${stIni} → ${stFim}`));
        children.push(kvDocx("Voto", voto || "—"));

        children.push(new Paragraph({ children: [new TextRun(" ")], spacing: { after: 160 } }));
      }
    }

    const footer = new Footer({
      children: [
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [
            new TextRun("Página "),
            new TextRun({ children: [PageNumber.CURRENT] }),
            new TextRun(" de "),
            new TextRun({ children: [PageNumber.TOTAL_PAGES] }),
          ],
        }),
      ],
    });

    const doc = new Document({
      sections: [{ properties: {}, footers: { default: footer }, children }],
    });

    const blob = await Packer.toBlob(doc);
    const filename = `${lastFilenameBase || "relatorio_pauta"}.docx`;
    saveAs(blob, filename);
    return filename;
  }

  function summaryParagraphsDocx(entries) {
    if (!entries.length) {
      return [new window.docx.Paragraph({ children: [new window.docx.TextRun({ text: "—" })], spacing: { after: 40 } })];
    }
    return entries.map(([k, n]) => new window.docx.Paragraph({ children: [new window.docx.TextRun({ text: `${k}: ${n}` })], spacing: { after: 40 } }));
  }

  function sectionDocx(title) {
    return new window.docx.Paragraph({
      children: [new window.docx.TextRun({ text: title, bold: true, color: "334155" })],
      spacing: { before: 120, after: 40 },
    });
  }

  function kvDocx(k, v) {
    return new window.docx.Paragraph({
      children: [new window.docx.TextRun({ text: `${k}: `, bold: true }), new window.docx.TextRun({ text: String(v ?? "") || "—" })],
      spacing: { after: 20 },
    });
  }

  function listDocx(label, arr) {
    if (!arr || arr.length === 0) return kvDocx(label, "—");
    const text = arr.join("\n");
    return new window.docx.Paragraph({
      children: [new window.docx.TextRun({ text: `${label}: `, bold: true }), new window.docx.TextRun({ text })],
      spacing: { after: 20 },
    });
  }

  function hexToDocxColor(hex) {
    return String(hex || "#111827").replace("#", "").toUpperCase();
  }

  function dataUrlToUint8Array(dataUrl) {
    const b64 = String(dataUrl).split(",")[1] || "";
    const bin = atob(b64);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return bytes;
  }

  // ===== Botões Gerar =====
  on(btnPdf, "click", async () => {
    const res = validateAndGetSessionRows();
    if (!res.ok) {
      userAlert(res.message || "Não foi possível gerar.");
      return;
    }

    userAlert("Gerando PDF...");
    try {
      const filename = await generatePdf(res.rows);
      userAlert(`PDF gerado: ${filename}`);
    } catch (err) {
      console.error(err);
      userAlert("Não foi possível gerar o PDF. Tente novamente.");
    }
  });

  on(btnDocx, "click", async () => {
    const res = validateAndGetSessionRows();
    if (!res.ok) {
      userAlert(res.message || "Não foi possível gerar.");
      return;
    }

    userAlert("Gerando DOCX...");
    try {
      const filename = await generateDocx(res.rows);
      userAlert(`DOCX gerado: ${filename}`);
    } catch (err) {
      console.error(err);
      userAlert("Não foi possível gerar o DOCX. Tente novamente.");
    }
  });

  // init
  updateButtons();

  return {
    destroy() {
      listeners.forEach((off) => off());
    },
  };
}
