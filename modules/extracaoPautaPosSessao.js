/* global XLSX */

import { safeFilename } from "../shared/helpers.js";

export function mount(container) {
  container.innerHTML = `
    <div class="card">
      <div class="card-body">
        <p class="mb-3">
          Envie os dois HTMLs da pauta (Página da Sessão e Imprimir Relação)
          para gerar a planilha consolidada com as informações dos processos em pauta na sessão (pós-sessão).
          </br></br>
          Para gerar o HTML 1 (Página da Sessão): </br></br>
          1 - Vá no Processo Eletrônico, menu Julgamento Colegiado - Sessões Plenárias - Plenário Presencial
          e selecione a sessão desejada.</br>
          2 - Entre na aba Relação de Julgamento </br>
          3 - Aperte Ctrl+S, selecione o tipo "Página da web, completa(*.htm;*html)" e salve o arquivo
         </br></br>
          Para gerar o HTML 2 (Imprimir Relação):</br></br>
          1 - Após salvar o HTML1, vá no final da página onde você está, clique no botão "Imprimir Relação" </br>
          2 - Na janela que abre, selecione HTML </br>
          3 - Vai abrir uma nova página e você deve salvá-la também através do Ctrl+S </br>
        </br>
        </p>

        <div class="alert alert-info mb-3" role="alert">
          A diferença entre a extração pré-sessão e a pós-sessão é que a planilha pós-sessão inclui a coluna “Status Final”, que apresenta o status do processo ao término da sessão.
        </div>

        <div class="row g-3">
          <div class="col-md-6">
            <label for="htmlDoc1" class="form-label">HTML 1 - Página da Sessão</label>
            <input id="htmlDoc1" class="form-control" type="file" accept=".html,text/html" />
          </div>

          <div class="col-md-6">
            <label for="htmlDoc2" class="form-label">HTML 2 - Imprimir Relação</label>
            <input id="htmlDoc2" class="form-control" type="file" accept=".html,text/html" />
          </div>
        </div>

        <div class="d-flex gap-2 mt-3">
          <button id="btnGerar" class="btn btn-primary" disabled>Gerar XLSX</button>
        </div>

        <div id="status" class="small text-muted mt-3">Selecione os dois arquivos HTML.</div>
      </div>
    </div>
  `;

  const htmlDoc1El = container.querySelector("#htmlDoc1");
  const htmlDoc2El = container.querySelector("#htmlDoc2");
  const btnGerarEl = container.querySelector("#btnGerar");
  const statusEl = container.querySelector("#status");

  let fileDoc1 = null;
  let fileDoc2 = null;

  const listeners = [];
  function on(el, evt, fn) {
    el.addEventListener(evt, fn);
    listeners.push(() => el.removeEventListener(evt, fn));
  }

  function setStatus(text) {
    statusEl.textContent = text;
  }

  function updateButtonState() {
    btnGerarEl.disabled = !(fileDoc1 && fileDoc2);
  }

  on(htmlDoc1El, "change", (event) => {
    fileDoc1 = event.target.files?.[0] || null;
    updateButtonState();
  });

  on(htmlDoc2El, "change", (event) => {
    fileDoc2 = event.target.files?.[0] || null;
    updateButtonState();
  });

  on(btnGerarEl, "click", async () => {
    if (!fileDoc1 || !fileDoc2) return;
    if (!window.XLSX) {
      setStatus("Biblioteca XLSX não carregada.");
      return;
    }

    try {
      setStatus("Lendo HTMLs...");

      const [html1, html2] = await Promise.all([fileDoc1.text(), fileDoc2.text()]);
      const processosDoc2 = extrairProcessosDoc2(html2);
      const resultado = extrairLinhasDoc1(html1, processosDoc2);

      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet(resultado.aoa);

      ws["!cols"] = [
        { wch: 48 },
        { wch: 24 },
        { wch: 24 },
        { wch: 26 },
        { wch: 36 },
        { wch: 18 },
        { wch: 24 },
        { wch: 45 },
        { wch: 42 },
        { wch: 20 },
        { wch: 14 },
      ];

      aplicarEstilos(ws, resultado);

      XLSX.utils.book_append_sheet(wb, ws, "Extração Pauta");
      const base = safeFilename((fileDoc1.name || "extracao-pauta").replace(/\.html?$/i, ""));
      XLSX.writeFile(wb, `${base}_extracao_pos_sessao.xlsx`);

      setStatus(
        `Planilha gerada com ${resultado.incluidos} processos. ` +
          `${resultado.adiados} processo(s) não incluído(s) por status Adiado.`,
      );
    } catch (error) {
      console.error(error);
      setStatus(`Erro: ${error.message || error}`);
    }
  });

  return {
    destroy() {
      listeners.forEach((off) => off());
      container.innerHTML = "";
    },
  };
}

function normalizarCabecalho(texto) {
  return String(texto || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toLowerCase();
}

function limparTexto(texto) {
  return String(texto || "").replace(/\s+/g, " ").trim();
}

function removerParenteses(texto) {
  return String(texto || "").replace(/[()]/g, "").trim();
}

function removerConteudoEntreParenteses(texto) {
  return String(texto || "")
    .replace(/\s*\([^)]*\)/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizarProcesso(valor) {
  const texto = limparTexto(valor).toUpperCase();
  const match = texto.match(/\d[\d./-]{5,}/);
  return match ? match[0] : texto;
}

function splitComQuebra(el) {
  return (el?.innerText || "")
    .split(/\r?\n/)
    .map((linha) => linha.trim())
    .filter(Boolean);
}

function mapearStatus(tipoInclusao) {
  const v = limparTexto(tipoInclusao);
  if (v === "Pauta") return "Pauta";
  if (v === "Extrapauta") return "Extrapauta";
  if (v === "Pedido de Vista devolvido") return "Devolução Vista";
  if (v === "Destacado de Sessão Virtual") return "Destaque Sessão Virtual";
  return "";
}

function deduplicarLinhas(valores) {
  const vistos = new Set();
  const unicos = [];

  valores.forEach((valor) => {
    const texto = limparTexto(valor);
    if (!texto) return;

    const chave = texto.toLocaleLowerCase("pt-BR");
    if (vistos.has(chave)) return;

    vistos.add(chave);
    unicos.push(texto);
  });

  return unicos;
}

function mapearVoto(celulaVoto) {
  const texto = limparTexto(celulaVoto?.textContent || "");
  if (/nao disponibilizado|não disponibilizado/i.test(texto)) return "Indisponível";
  if (celulaVoto?.querySelector("button")) return "Listado";
  return "";
}

function extrairProcessosDoc2(html) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(html, "text/html");
  const processos = new Map();

  const linhas = doc.querySelectorAll("tr");
  linhas.forEach((tr) => {
    const cols = tr.querySelectorAll("td");
    if (cols.length < 2) return;

    const processoRaw = limparTexto(cols[0].textContent || "");
    if (!processoRaw) return;

    const processo = normalizarProcesso(processoRaw);
    if (!processo) return;

    const linhasCol2 = splitComQuebra(cols[1]);
    if (!linhasCol2.length) {
      processos.set(processo, { interessados: "", advogados: "" });
      return;
    }

    const interessados = [];
    const advogados = [];

    linhasCol2.slice(1).forEach((linha) => {
      const linhaLimpa = limparTexto(linha);
      if (!linhaLimpa) return;

      if (/\badv\.?\b/i.test(linhaLimpa)) {
        const advogadoLimpo = removerParenteses(linhaLimpa)
          .replace(/^adv\.?\s*:?\s*/i, "")
          .replace(/^[-:;,.]+\s*/, "")
          .trim();
        if (advogadoLimpo) advogados.push(advogadoLimpo);
        return;
      }

      interessados.push(removerParenteses(linhaLimpa));
    });

    processos.set(processo, {
      interessados: deduplicarLinhas(interessados).join("\n"),
      advogados: deduplicarLinhas(advogados).join("\n"),
    });
  });

  return processos;
}

function localizarTabelaRelacao(doc) {
  const tabelas = Array.from(doc.querySelectorAll("table"));
  for (const tabela of tabelas) {
    const headers = Array.from(tabela.querySelectorAll("tr th, tr td"))
      .slice(0, 12)
      .map((cell) => normalizarCabecalho(cell.textContent));

    if (headers.some((h) => h.includes("processo")) && headers.some((h) => h.includes("tipo de inclusao"))) {
      return tabela;
    }
  }
  return null;
}

function extrairLinhasDoc1(html, processosDoc2) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(html, "text/html");
  const tabela = localizarTabelaRelacao(doc);

  if (!tabela) {
    throw new Error('Tabela "Página da Sessão" não encontrada no HTML 1.');
  }

  const cabecalho = [
    "Processo",
    "Status",
    "Status Final",
    "Relator",
    "Órgão",
    "Modalidade",
    "Tipo Processo",
    "Interessados",
    "Advogados",
    "Sistema de Tramitação",
    "Voto",
  ];

  const aoa = [cabecalho];
  const statusNaoPautaRows = [];
  let adiados = 0;
  let incluidos = 0;

  const linhas = Array.from(tabela.querySelectorAll("tr"));
  const headerRow = linhas.find((tr) => tr.querySelector("th")) || linhas[0];
  const headerCells = Array.from(headerRow.querySelectorAll("th,td"));
  const idx = {
    voto: -1,
    modalidade: -1,
    tipo: -1,
    unidade: -1,
    processo: -1,
    tipoInclusao: -1,
    relator: -1,
    situacao: -1,
  };

  headerCells.forEach((cell, i) => {
    const h = normalizarCabecalho(cell.textContent);
    if (h === "voto") idx.voto = i;
    if (h === "modalidade") idx.modalidade = i;
    if (h === "tipo") idx.tipo = i;
    if (h.includes("unidade jurisdicionada")) idx.unidade = i;
    if (h === "processo") idx.processo = i;
    if (h.includes("tipo de inclusao")) idx.tipoInclusao = i;
    if (h === "relator") idx.relator = i;
    if (h === "situacao") idx.situacao = i;
  });

  linhas.forEach((tr) => {
    if (tr === headerRow || tr.querySelector("th")) return;
    const tds = tr.querySelectorAll("td");
    if (!tds.length) return;

    const tipoInclusao = limparTexto(tds[idx.tipoInclusao]?.textContent || "");
    if (tipoInclusao === "Adiado") {
      adiados += 1;
      return;
    }

    const processoTexto = limparTexto(tds[idx.processo]?.textContent || "");
    if (!processoTexto) return;

    const vinculadoMatch = processoTexto.match(/\(?\s*(VINCULADO AO CONSELHEIRO\s+.+?)\s*\)?$/i);
    const processoNumero = vinculadoMatch
      ? limparTexto(processoTexto.replace(vinculadoMatch[0], ""))
      : processoTexto;

    const processoKey = normalizarProcesso(processoNumero);
    const detalhes = processosDoc2.get(processoKey) || { interessados: "", advogados: "" };

    const processoFinal = vinculadoMatch
      ? `${processoNumero}\n⚠️ ${removerParenteses(vinculadoMatch[1])}`
      : processoNumero;

    const statusFinal = mapearStatus(tipoInclusao);
    const idxSituacao = idx.situacao >= 0 ? idx.situacao : 8;
    const situacaoFinal = removerConteudoEntreParenteses(tds[idxSituacao]?.textContent || "");
    const rowIndex = aoa.length + 1;

    if (statusFinal && statusFinal !== "Pauta") {
      statusNaoPautaRows.push(rowIndex);
    }

    aoa.push([
      processoFinal,
      statusFinal,
      situacaoFinal,
      limparTexto(tds[idx.relator]?.textContent || ""),
      limparTexto(tds[idx.unidade]?.textContent || ""),
      limparTexto(tds[idx.modalidade]?.textContent || ""),
      limparTexto(tds[idx.tipo]?.textContent || ""),
      detalhes.interessados,
      detalhes.advogados,
      "E-TCE",
      mapearVoto(tds[idx.voto]),
    ]);

    incluidos += 1;
  });

  return { aoa, adiados, incluidos, statusNaoPautaRows };
}

function aplicarEstilos(ws, resultado) {
  const range = XLSX.utils.decode_range(ws["!ref"]);

  for (let row = 0; row <= range.e.r; row++) {
    for (let col = 0; col <= range.e.c; col++) {
      const addr = XLSX.utils.encode_cell({ r: row, c: col });
      const cell = ws[addr];
      if (!cell) continue;

      cell.s = {
        alignment: { vertical: "top", wrapText: true },
        font: { name: "Calibri", sz: row === 0 ? 12 : 11, bold: row === 0 },
      };
    }
  }

  resultado.statusNaoPautaRows.forEach((excelRow) => {
    const statusCell = ws[`B${excelRow}`];
    if (!statusCell) return;
    statusCell.s = {
      ...(statusCell.s || {}),
      fill: { fgColor: { rgb: "FFF8CBAD" } },
      alignment: { vertical: "top", wrapText: true },
    };
  });

}
