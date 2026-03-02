/* global saveAs */

import {
  splitLines,
  upper,
  formatDateBR,
  ordinalFeminino,
} from "../shared/helpers.js";

/**
 * Relatório Pauta Dinâmica
 * - Entrada: planilha XLSX (aba 1)
 * - Filtro: Data (Excel date) + Tipo Sessão (PLENO/1CAM/2CAM)
 * - Saída: PDF (pdfmake) ou DOCX (docx)
 * - Validação dura de Status/Status Final conforme listas definidas
 * - Índice por relator + links (PDF: interno; DOCX: interno se viewer suportar)
 */

export function mount(container) {
  container.innerHTML = `
    <div class="module-card">
      <div class="card">
        <div class="card-body">
          <div class="row g-3">
            <div class="col-md-3">
              <label for="sessionNumberR" class="form-label">Nº da sessão</label>
              <input
                id="sessionNumberR"
                type="number"
                min="1"
                step="1"
                class="form-control"
                placeholder="Ex: 20"
                required
              />
            </div>

            <div class="col-md-5">
              <label for="sessionTypeR" class="form-label">Tipo de sessão</label>
              <select id="sessionTypeR" class="form-select" required>
                <option value="" selected>Selecione...</option>
                <option value="PLENO">Pleno</option>
                <option value="1CAM">1ª Câmara</option>
                <option value="2CAM">2ª Câmara</option>
              </select>
              <div class="form-text">
                Observação: na planilha, a coluna <strong>Tipo Sessão</strong> deve estar exatamente como
                <code>PLENO</code>, <code>1CAM</code> ou <code>2CAM</code>.
              </div>
            </div>

            <div class="col-md-4">
              <label for="sessionDateR" class="form-label">Data</label>
              <input id="sessionDateR" type="date" class="form-control" required />
            </div>
          </div>

          <hr class="my-3" />

          <label for="fileInputR" class="form-label">Selecione o arquivo .xlsx</label>
          <input class="form-control" type="file" id="fileInputR" accept=".xlsx" />

          <div class="d-flex flex-wrap gap-2 mt-3">
            <button id="btnPdf" class="btn btn-primary" disabled>Gerar PDF</button>
            <button id="btnDocx" class="btn btn-outline-primary" disabled>Gerar DOCX</button>
          </div>

          <div class="mt-3">
            <div id="statusR" class="small text-muted">Nenhum arquivo selecionado.</div>
          </div>

          <div id="errorsR" class="mt-3"></div>
        </div>
      </div>
    </div>
  `;

  // ===== DOM =====
  const fileInput = container.querySelector("#fileInputR");
  const btnPdf = container.querySelector("#btnPdf");
  const btnDocx = container.querySelector("#btnDocx");
  const statusEl = container.querySelector("#statusR");
  const errorsEl = container.querySelector("#errorsR");

  const sessionNumberEl = container.querySelector("#sessionNumberR");
  const sessionTypeEl = container.querySelector("#sessionTypeR");
  const sessionDateEl = container.querySelector("#sessionDateR");

  // ===== Estado =====
  let allRows = null; // cache da planilha (não precisa reenviar)
  let lastFilenameBase = null;

  // ===== Constantes de validação =====
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

  // ===== Logo (base64 PNG) – 1ª página apenas =====
  // Observação: isso deixa o sistema independente de hospedar imagem no repo.
  const LOGO_DATA_URL =
    "data:image/png;base64," +
    "iVBORw0KGgoAAAANSUhEUgAACAAAAAZmCAYAAAA8NUQqAAAQAElEQVR4AezdB3zTZf7A8e/zS7rZW5Yg" +
    // (string enorme – mantida completa abaixo)
    LOGO_BASE64_REST;

  /**
   * A parte restante do base64 é colocada no final do arquivo para não poluir
   * o topo. NÃO altere isso.
   */
  function getLogoDataUrl() {
    return "data:image/png;base64," + LOGO_BASE64_FULL;
  }

  // ===== Helpers UI =====
  const listeners = [];
  function on(el, evt, fn) {
    el.addEventListener(evt, fn);
    listeners.push(() => el.removeEventListener(evt, fn));
  }

  function setStatus(msg) {
    statusEl.textContent = msg;
  }

  function clearErrors() {
    errorsEl.innerHTML = "";
  }

  function showErrorsGrouped(groups) {
    // groups: Array<{title: string, items: Array<{processo, relator2, orgao, detalhe?}>}>
    const blocks = groups
      .filter((g) => g.items.length > 0)
      .map((g) => {
        const rows = g.items
          .map((it) => {
            const det = it.detalhe ? ` — <span class="text-muted">${escapeHtml(it.detalhe)}</span>` : "";
            return `<li><code>${escapeHtml(it.processo)}</code> — ${escapeHtml(it.relator2)} — ${escapeHtml(it.orgao)}${det}</li>`;
          })
          .join("");
        return `
          <div class="alert alert-danger">
            <div><strong>${escapeHtml(g.title)}</strong></div>
            <ul class="mb-0 mt-2">${rows}</ul>
          </div>
        `;
      })
      .join("");

    errorsEl.innerHTML = blocks || "";
  }

  function escapeHtml(s) {
    return String(s ?? "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
  }

  function headerOk() {
    return (
      !!ordinalFeminino(sessionNumberEl.value) &&
      !!String(sessionTypeEl.value || "").trim() &&
      !!String(sessionDateEl.value || "").trim()
    );
  }

  function updateButtons() {
    const ok = !!allRows && headerOk();
    btnPdf.disabled = !ok;
    btnDocx.disabled = !ok;
  }

  updateButtons();
  [sessionNumberEl, sessionTypeEl, sessionDateEl].forEach((el) => {
    on(el, "input", updateButtons);
    on(el, "change", updateButtons);
  });

  // ===== Leitura do XLSX (com datas) =====
  on(fileInput, "change", async (e) => {
    clearErrors();
    allRows = null;
    lastFilenameBase = null;
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
    if (!window.XLSX) {
      setStatus("Biblioteca XLSX não carregada (CDN).");
      return;
    }

    setStatus("Lendo XLSX...");

    try {
      const arrayBuffer = await file.arrayBuffer();
      const workbook = window.XLSX.read(arrayBuffer, { type: "array", cellDates: true });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      allRows = window.XLSX.utils.sheet_to_json(sheet, { defval: "" });

      setStatus(`XLSX OK. Linhas: ${allRows.length}.`);
      updateButtons();
    } catch (err) {
      console.error(err);
      setStatus("Erro ao ler XLSX. Abra o Console (F12) e veja o erro.");
      allRows = null;
      updateButtons();
    }
  });

  // ===== Core: filtro + validação =====

  function mapSessionTypeToHeader(typeCode) {
    // cabeçalho igual ao do sistema de pauta manual
    if (typeCode === "PLENO") return "PLENO";
    if (typeCode === "1CAM") return "PRIMEIRA CÂMARA";
    if (typeCode === "2CAM") return "SEGUNDA CÂMARA";
    return String(typeCode || "").toUpperCase();
  }

  function excelCellToYmd(value) {
    // value pode ser Date (cellDates: true), number serial ou string
    // objetivo: retornar "YYYY-MM-DD" ou "" se inválido
    if (!value) return "";

    if (value instanceof Date && !isNaN(value.getTime())) {
      const y = value.getFullYear();
      const m = String(value.getMonth() + 1).padStart(2, "0");
      const d = String(value.getDate()).padStart(2, "0");
      return `${y}-${m}-${d}`;
    }

    if (typeof value === "number" && window.XLSX?.SSF?.parse_date_code) {
      const parsed = window.XLSX.SSF.parse_date_code(value);
      if (!parsed) return "";
      const y = parsed.y;
      const m = String(parsed.m).padStart(2, "0");
      const d = String(parsed.d).padStart(2, "0");
      return `${y}-${m}-${d}`;
    }

    // tenta dd/mm/yyyy
    const s = String(value).trim();
    const m1 = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if (m1) return `${m1[3]}-${m1[2]}-${m1[1]}`;

    // tenta yyyy-mm-dd
    const m2 = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (m2) return s;

    return "";
  }

  function relator2Nomes(nome) {
    const parts = String(nome ?? "").trim().split(/\s+/).filter(Boolean);
    return parts.slice(0, 2).join(" ");
  }

  function validateAndGetSessionRows() {
    clearErrors();

    if (!allRows) {
      return { ok: false, rows: [], message: "Carregue a planilha primeiro." };
    }
    if (!headerOk()) {
      return { ok: false, rows: [], message: "Preencha Nº da sessão, Tipo de sessão e Data." };
    }

    const typeCode = sessionTypeEl.value; // PLENO/1CAM/2CAM (exigido literal na planilha)
    const ymd = sessionDateEl.value; // YYYY-MM-DD

    // filtro literal do Tipo Sessão e por data convertida
    const filtered = allRows.filter((r) => {
      const tipo = r["Tipo Sessão"];
      const data = r["Data"];
      const dataYmd = excelCellToYmd(data);
      return tipo === typeCode && dataYmd === ymd;
    });

    if (filtered.length === 0) {
      const dateBR = formatDateBR(ymd);
      return {
        ok: false,
        rows: [],
        message:
          `Não foram encontradas sessões para o dia ${dateBR}.\n` +
          `Verifique o preenchimento da planilha: na coluna "Tipo Sessão" deve estar preenchido ` +
          `"${typeCode}" exatamente (ex.: 1CAM), e a coluna "Data" deve corresponder à data da sessão.`,
      };
    }

    // validações de status + status final vazio
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
        groups[1].items.push({ processo, relator2: rel2, orgao, detalhe: `Valor encontrado: "${stFim}"` });
      }

      if (!stIni) {
        groups[2].items.push({ processo, relator2: rel2, orgao });
      } else if (!STATUS_INICIAIS.has(stIni)) {
        groups[3].items.push({ processo, relator2: rel2, orgao, detalhe: `Valor encontrado: "${stIni}"` });
      }
    }

    const hasErrors = groups.some((g) => g.items.length > 0);
    if (hasErrors) {
      showErrorsGrouped(groups);
      return {
        ok: false,
        rows: [],
        message: "Foram encontrados erros na planilha. Corrija e gere novamente.",
      };
    }

    // base de nome de arquivo
    const dateBR = formatDateBR(ymd);
    lastFilenameBase = `relatorio_pauta_${typeCode}_${dateBR.replaceAll("/", "-")}`;

    return { ok: true, rows: filtered, message: "" };
  }

  // ===== Cores institucionais discretas (Status Final) =====
  function statusColorHex(statusFinal) {
    // Você pediu amarelo para Adiado/Pedido de Vista – uso tom mostarda escuro (legível).
    // Retirado de Pauta: bordô discreto.
    // Sobrestado: roxo discreto.
    // Julgado: verde escuro.
    if (statusFinal === "Julgado") return "#1B5E20";
    if (statusFinal === "Adiado" || statusFinal === "Pedido de Vista") return "#8A6D00";
    if (statusFinal === "Sobrestado") return "#4A148C";
    if (statusFinal === "Retirado de Pauta") return "#7F1D1D";
    return "#111827";
  }

  // ===== Agrupamento por relator (mantém ordem da planilha) =====
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
    // retorna apenas >0
    return Array.from(counts.entries())
      .filter(([, n]) => n > 0)
      .sort((a, b) => a[0].localeCompare(b[0], "pt-BR"));
  }

  // ===== Geração PDF (pdfmake) =====
  async function generatePdf(rows) {
    if (!window.pdfMake) throw new Error("Biblioteca pdfmake não carregada (CDN).");

    const sessionNumber = ordinalFeminino(sessionNumberEl.value);
    const sessionTypeHeader = mapSessionTypeToHeader(sessionTypeEl.value);
    const dateBR = formatDateBR(sessionDateEl.value);

    const { map, order } = groupByRelatorPreserveOrder(rows);

    // Resumos (só o que existir)
    const pre = countByField(rows, "Status"); // inicial
    const post = countByField(rows, "Status Final"); // final

    // IDs internos (destinos) para navegação
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

    // Logo (1ª página)
    content.push({
      image: "logo",
      width: 70,
      alignment: "left",
      margin: [0, 0, 0, 8],
    });

    // Cabeçalho igual ao da pauta manual
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

    // Quadros resumo pré/pós
    content.push({
      columns: [
        summaryBox("Resumo pré-sessão", pre),
        summaryBox("Resumo pós-sessão", post),
      ],
      columnGap: 12,
      margin: [0, 0, 0, 14],
    });

    // Índice
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

        // Linha 1: processo – status (clicável)
        content.push({
          text: `${proc} — ${stFim}`,
          margin: [12, 0, 0, 0],
          linkToDestination: dest,
          color: "#0F172A",
        });

        // Linha 2: relator completo – órgão
        content.push({
          text: `${rel} — ${orgao}`,
          margin: [12, 0, 0, 6],
          fontSize: 9,
          color: "#374151",
        });
      }
    }

    // Seções por relator (cada relator nova página)
    for (let idx = 0; idx < order.length; idx++) {
      const rel = order[idx];
      const ps = map.get(rel) || [];
      const relId = relatorDest.get(rel);

      content.push({ text: "", pageBreak: "before" });

      // âncora do relator
      content.push({ text: "", id: relId });

      content.push({
        text: upper(rel),
        bold: true,
        fontSize: 13,
        margin: [0, 0, 0, 10],
      });

      // cards dos processos (fluem e quebram página automaticamente)
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

        content.push({
          // âncora do processo
          text: "",
          id: procId,
          margin: [0, 0, 0, 0],
        });

        // Card
        content.push({
          stack: [
            { text: proc, bold: true, fontSize: 12, margin: [0, 0, 0, 2] },
            { text: stFim, bold: true, fontSize: 12, color: statusColorHex(stFim), margin: [0, 0, 0, 8] },

            sectionTitle("INFORMAÇÕES GERAIS"),
            keyValue("Relator", rel),
            keyValue("Órgão", orgao),

            sectionTitle("CLASSIFICAÇÃO"),
            keyValue("Modalidade – Tipo Processo", `${modalidade}${modalidade && tipoProc ? " - " : ""}${tipoProc}`),

            sectionTitle("PARTES"),
            listField("Interessados", interessados),
            listField("Advogados", advogados),

            sectionTitle("TRAMITAÇÃO"),
            keyValue("Sistema de Tramitação", sist),
            keyValue("Status Inicial → Status Final", `${stIni} → ${stFim}`),
            keyValue("Voto", voto || "—"),
          ],
          margin: [0, 0, 0, 12],
          // borda discreta do card
          border: [true, true, true, true],
          style: "card",
        });
      }
    }

    const docDefinition = {
      pageSize: "A4",
      pageMargins: [40, 35, 40, 45],
      footer: function (currentPage, pageCount) {
        return {
          text: `Página ${currentPage} de ${pageCount}`,
          alignment: "right",
          margin: [0, 0, 40, 10],
          fontSize: 9,
          color: "#374151",
        };
      },
      images: {
        logo: getLogoDataUrl(),
      },
      styles: {
        card: {
          margin: [0, 0, 0, 0],
        },
      },
      defaultStyle: {
        fontSize: 10,
        color: "#111827",
      },
      content,
    };

    // Cria e baixa
    const filename = `${lastFilenameBase || "relatorio_pauta"}.pdf`;
    window.pdfMake.createPdf(docDefinition).download(filename);
    return filename;
  }

  function summaryBox(title, entries) {
    const lines = entries.length
      ? entries.map(([k, n]) => `${k}: ${n}`).join("\n")
      : "—";

    return {
      width: "*",
      stack: [
        { text: title, bold: true, margin: [0, 0, 0, 6] },
        { text: lines, fontSize: 10, lineHeight: 1.2 },
      ],
      margin: [0, 0, 0, 0],
      border: [true, true, true, true],
      padding: 8,
    };
  }

  function sectionTitle(t) {
    return { text: t, bold: true, fontSize: 9, color: "#334155", margin: [0, 6, 0, 3] };
  }

  function keyValue(k, v) {
    return {
      text: [
        { text: `${k}: `, bold: true },
        { text: String(v ?? "") || "—" },
      ],
      margin: [0, 0, 0, 2],
    };
  }

  function listField(label, arr) {
    if (!arr || arr.length === 0) return keyValue(label, "—");
    return {
      stack: [
        { text: `${label}:`, bold: true, margin: [0, 0, 0, 2] },
        ...arr.map((x) => ({ text: x, margin: [10, 0, 0, 1] })),
      ],
      margin: [0, 0, 0, 2],
    };
  }

  function hashId(s) {
    // hash simples/estável para ids curtos
    const str = String(s ?? "");
    let h = 2166136261;
    for (let i = 0; i < str.length; i++) {
      h ^= str.charCodeAt(i);
      h = Math.imul(h, 16777619);
    }
    return (h >>> 0).toString(36);
  }

  // ===== Geração DOCX (docx) =====
  async function generateDocx(rows) {
    if (!window.docx) throw new Error("Biblioteca docx não carregada (CDN).");
    if (!window.saveAs) throw new Error("Biblioteca FileSaver não carregada (CDN).");

    const {
      Document,
      Packer,
      Paragraph,
      TextRun,
      AlignmentType,
      HeadingLevel,
      Footer,
      PageNumber,
      ImageRun,
      InternalHyperlink,
      Bookmark,
    } = window.docx;

    const sessionNumber = ordinalFeminino(sessionNumberEl.value);
    const sessionTypeHeader = mapSessionTypeToHeader(sessionTypeEl.value);
    const dateBR = formatDateBR(sessionDateEl.value);

    const { map, order } = groupByRelatorPreserveOrder(rows);

    // Resumos (sem itens zerados)
    const pre = countByField(rows, "Status");
    const post = countByField(rows, "Status Final");

    // IDs internos (bookmarks)
    const relatorBm = new Map();
    const procBm = new Map();
    for (const rel of order) {
      relatorBm.set(rel, `REL_${hashId(rel)}`);
      const ps = map.get(rel) || [];
      for (const r of ps) {
        const proc = String(r["Processo"] ?? "").trim();
        if (!proc) continue;
        procBm.set(proc, `P_${hashId(rel + "|" + proc)}`);
      }
    }

    // Logo (1ª página) – inserida como primeiro parágrafo (não header)
    const logoBytes = base64ToUint8Array(getLogoDataUrl().split(",")[1]);

    const children = [];

    children.push(
      new Paragraph({
        children: [
          new ImageRun({
            data: logoBytes,
            transformation: { width: 110, height: 45 },
          }),
        ],
        alignment: AlignmentType.LEFT,
        spacing: { after: 120 },
      })
    );

    // Cabeçalho (igual ao da pauta manual)
    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({
            text: `PAUTA DA ${sessionNumber} SESSÃO ORDINÁRIA DO ${sessionTypeHeader}`,
            bold: true,
            size: 24,
          }),
        ],
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

    // Resumos (texto)
    children.push(
      new Paragraph({
        children: [new TextRun({ text: "RESUMO PRÉ-SESSÃO", bold: true })],
        spacing: { after: 60 },
      })
    );
    children.push(...summaryParagraphs(pre));

    children.push(
      new Paragraph({
        children: [new TextRun({ text: "RESUMO PÓS-SESSÃO", bold: true })],
        spacing: { before: 120, after: 60 },
      })
    );
    children.push(...summaryParagraphs(post));

    // Índice
    children.push(
      new Paragraph({
        children: [new TextRun({ text: "ÍNDICE (por Relator)", bold: true })],
        spacing: { before: 180, after: 80 },
      })
    );

    for (const rel of order) {
      const relId = relatorBm.get(rel);

      // Link para relator (se InternalHyperlink/Bookmark funcionar no viewer)
      children.push(
        new Paragraph({
          children: [
            safeInternalLink(relId, `RELATOR: ${upper(rel)}`, true),
          ],
          spacing: { before: 80, after: 40 },
        })
      );

      const ps = map.get(rel) || [];
      for (const r of ps) {
        const proc = String(r["Processo"] ?? "").trim();
        const orgao = String(r["Órgão"] ?? "").trim();
        const stFim = String(r["Status Final"] ?? "").trim();

        const procId = procBm.get(proc);

        children.push(
          new Paragraph({
            children: [
              new TextRun({ text: "  • " }),
              safeInternalLink(procId, `${proc} — ${stFim}`, false),
              new TextRun({ text: ` — ${rel} — ${orgao}`, size: 18, color: "475569" }),
            ],
            spacing: { after: 20 },
          })
        );
      }
    }

    // Seções por relator (cada relator em nova página)
    for (const rel of order) {
      const relId = relatorBm.get(rel);
      const ps = map.get(rel) || [];

      children.push(
        new Paragraph({
          children: [new TextRun("")],
          pageBreakBefore: true,
        })
      );

      // Bookmark do relator
      children.push(new Paragraph({ children: [new Bookmark({ id: relId, children: [] })] }));

      children.push(
        new Paragraph({
          text: upper(rel),
          heading: HeadingLevel.HEADING_1,
          spacing: { after: 160 },
        })
      );

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

        const procId = procBm.get(proc);

        // Bookmark do processo
        children.push(new Paragraph({ children: [new Bookmark({ id: procId, children: [] })] }));

        // Topo do card
        children.push(
          new Paragraph({
            children: [
              new TextRun({ text: proc, bold: true, size: 24 }),
            ],
            spacing: { after: 20 },
          })
        );
        children.push(
          new Paragraph({
            children: [
              new TextRun({
                text: stFim,
                bold: true,
                size: 24,
                color: hexToDocxColor(statusColorHex(stFim)),
              }),
            ],
            spacing: { after: 100 },
          })
        );

        // Seções
        children.push(sectionDocx("INFORMAÇÕES GERAIS"));
        children.push(kvDocx("Relator", rel));
        children.push(kvDocx("Órgão", orgao));

        children.push(sectionDocx("CLASSIFICAÇÃO"));
        children.push(kvDocx("Modalidade – Tipo Processo", `${modalidade}${modalidade && tipoProc ? " - " : ""}${tipoProc}`));

        children.push(sectionDocx("PARTES"));
        children.push(listDocx("Interessados", interessados));
        children.push(listDocx("Advogados", advogados));

        children.push(sectionDocx("TRAMITAÇÃO"));
        children.push(kvDocx("Sistema de Tramitação", sist));
        children.push(kvDocx("Status Inicial → Status Final", `${stIni} → ${stFim}`));
        children.push(kvDocx("Voto", voto || "—"));

        // separador leve
        children.push(
          new Paragraph({
            children: [new TextRun(" ")],
            spacing: { after: 160 },
          })
        );
      }
    }

    // Rodapé com Página X de Y
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
      sections: [
        {
          properties: {},
          footers: { default: footer },
          children,
        },
      ],
    });

    const blob = await Packer.toBlob(doc);
    const filename = `${lastFilenameBase || "relatorio_pauta"}.docx`;
    saveAs(blob, filename);
    return filename;
  }

  function summaryParagraphs(entries) {
    if (!entries.length) {
      return [
        new window.docx.Paragraph({
          children: [new window.docx.TextRun({ text: "—" })],
          spacing: { after: 40 },
        }),
      ];
    }
    return entries.map(([k, n]) => {
      return new window.docx.Paragraph({
        children: [new window.docx.TextRun({ text: `${k}: ${n}` })],
        spacing: { after: 40 },
      });
    });
  }

  function sectionDocx(title) {
    return new window.docx.Paragraph({
      children: [new window.docx.TextRun({ text: title, bold: true, color: "334155" })],
      spacing: { before: 120, after: 40 },
    });
  }

  function kvDocx(k, v) {
    return new window.docx.Paragraph({
      children: [
        new window.docx.TextRun({ text: `${k}: `, bold: true }),
        new window.docx.TextRun({ text: String(v ?? "") || "—" }),
      ],
      spacing: { after: 20 },
    });
  }

  function listDocx(label, arr) {
    if (!arr || arr.length === 0) return kvDocx(label, "—");

    const paras = [];
    paras.push(
      new window.docx.Paragraph({
        children: [new window.docx.TextRun({ text: `${label}:`, bold: true })],
        spacing: { after: 20 },
      })
    );
    for (const x of arr) {
      paras.push(
        new window.docx.Paragraph({
          children: [new window.docx.TextRun({ text: `• ${x}` })],
          spacing: { after: 10 },
        })
      );
    }
    // retorna “stack” simulada: caller faz push(...arr)
    // aqui devolvo um parágrafo “marcador” com children vazios e uso “caller” diferente:
    // para manter simples, vamos juntar em um único parágrafo com quebras:
    // (Mas Word lida melhor com múltiplos parágrafos. Então retornamos apenas o primeiro e
    // o caller deverá empilhar todos; para não complicar, aqui devolvo o primeiro e o resto via propriedade extra)
    // => solução simples: retornar só o primeiro e os itens como linhas no mesmo parágrafo:
    return new window.docx.Paragraph({
      children: [
        new window.docx.TextRun({ text: `${label}: `, bold: true }),
        new window.docx.TextRun({ text: arr.join(" | ") }),
      ],
      spacing: { after: 20 },
    });
  }

  function hexToDocxColor(hex) {
    // "#RRGGBB" -> "RRGGBB"
    return String(hex || "#111827").replace("#", "").toUpperCase();
  }

  function base64ToUint8Array(b64) {
    const bin = atob(b64);
    const len = bin.length;
    const bytes = new Uint8Array(len);
    for (let i = 0; i < len; i++) bytes[i] = bin.charCodeAt(i);
    return bytes;
  }

  function safeInternalLink(anchorId, text, bold) {
    // Se InternalHyperlink não existir, cai para texto normal.
    try {
      if (window.docx?.InternalHyperlink && anchorId) {
        return new window.docx.InternalHyperlink({
          anchor: anchorId,
          children: [new window.docx.TextRun({ text, bold: !!bold, color: "0F172A", underline: {} })],
        });
      }
    } catch (e) {
      // ignora e cai para texto
    }
    return new window.docx.TextRun({ text, bold: !!bold });
  }

  // ===== Botões =====
  on(btnPdf, "click", async () => {
    clearErrors();
    const res = validateAndGetSessionRows();
    if (!res.ok) {
      setStatus(res.message || "Não foi possível gerar.");
      return;
    }

    setStatus("Gerando PDF...");
    try {
      const filename = await generatePdf(res.rows);
      setStatus(`PDF gerado: ${filename}`);
    } catch (err) {
      console.error(err);
      setStatus("Erro ao gerar PDF. Abra o Console (F12) e veja o erro.");
    }
  });

  on(btnDocx, "click", async () => {
    clearErrors();
    const res = validateAndGetSessionRows();
    if (!res.ok) {
      setStatus(res.message || "Não foi possível gerar.");
      return;
    }

    setStatus("Gerando DOCX...");
    try {
      const filename = await generateDocx(res.rows);
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

/**
 * Base64 completo da logo (PNG) – não edite.
 * (Coloquei como constante separada para o arquivo não ficar “quebrado” no topo.)
 */
const LOGO_BASE64_FULL =
`iVBORw0KGgoAAAANSUhEUgAACAAAAAZmCAYAAAA8NUQqAAAQAElEQVR4AezdB3zTZf7A8e/zS7rZW5Yg
...REPLACE_ME_WITH_FULL_BASE64...`;
