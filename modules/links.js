export function mount(container) {
  container.innerHTML = `
    <div class="module-card">
      <div class="card">
        <div class="card-body">
          <p class="text-muted small mb-3">
            Cole abaixo duas colunas copiadas do Google Sheets: na primeira, o número do processo; na segunda, o link.
          </p>

          <label for="linksInput" class="form-label">Entrada (2 colunas)</label>
          <textarea
            id="linksInput"
            class="form-control"
            rows="10"
            placeholder="0001234-56.2024.8.17.0000\thttps://..."
          ></textarea>

          <div class="d-flex flex-wrap gap-2 mt-3">
            <button id="btnGenerateLinks" class="btn btn-primary">Gerar coluna com link</button>
            <button id="btnCopyLinks" class="btn btn-outline-secondary" disabled>Copiar tudo</button>
          </div>

          <div class="mt-3">
            <div id="linksStatus" class="small text-muted">Aguardando dados.</div>
          </div>

          <div class="mt-3">
            <label for="linksOutput" class="form-label">Saída (1 coluna para colar no Sheets)</label>
            <textarea id="linksOutput" class="form-control" rows="10" readonly></textarea>
          </div>
        </div>
      </div>
    </div>
  `;

  const inputEl = container.querySelector("#linksInput");
  const outputEl = container.querySelector("#linksOutput");
  const statusEl = container.querySelector("#linksStatus");
  const btnGenerate = container.querySelector("#btnGenerateLinks");
  const btnCopy = container.querySelector("#btnCopyLinks");

  const listeners = [];

  function on(el, evt, fn) {
    el.addEventListener(evt, fn);
    listeners.push(() => el.removeEventListener(evt, fn));
  }

  function setStatus(msg) {
    statusEl.textContent = msg;
  }

  function parseRows(raw) {
    return raw
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter(Boolean)
      .map((line) => line.split("\t"));
  }

  function toHyperlinkFormula(processNumber, link) {
    const sanitizedProcess = String(processNumber || "").replace(/"/g, '""');
    const sanitizedLink = String(link || "").replace(/"/g, '""');
    return `=HYPERLINK("${sanitizedLink}";"${sanitizedProcess}")`;
  }

  on(btnGenerate, "click", () => {
    const input = inputEl.value || "";

    if (!input.trim()) {
      outputEl.value = "";
      btnCopy.disabled = true;
      setStatus("Cole os dados antes de gerar a coluna.");
      return;
    }

    const rows = parseRows(input);
    const output = [];
    let ignored = 0;

    for (const cols of rows) {
      const processNumber = (cols[0] || "").trim();
      const link = (cols[1] || "").trim();

      if (!processNumber || !link) {
        ignored += 1;
        continue;
      }

      output.push(toHyperlinkFormula(processNumber, link));
    }

    outputEl.value = output.join("\n");
    btnCopy.disabled = output.length === 0;

    if (output.length === 0) {
      setStatus("Nenhuma linha válida encontrada. Verifique se há 2 colunas (processo + link).");
      return;
    }

    const ignoredText = ignored > 0 ? ` (${ignored} linha(s) ignorada(s))` : "";
    setStatus(`Coluna gerada com ${output.length} linha(s).${ignoredText}`);
  });

  on(btnCopy, "click", async () => {
    if (!outputEl.value.trim()) return;

    try {
      await navigator.clipboard.writeText(outputEl.value);
      setStatus("Coluna copiada. Agora é só colar no Google Sheets.");
    } catch (_err) {
      outputEl.focus();
      outputEl.select();
      setStatus("Não foi possível copiar automaticamente. Use Ctrl+C na área de saída.");
    }
  });

  return {
    destroy() {
      listeners.forEach((off) => off());
    },
  };
}
