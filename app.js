import { clearElement } from "./shared/helpers.js";

const docTypeEl = document.getElementById("docType");
const container = document.getElementById("moduleContainer");

let currentModule = null;

function showMessage(html) {
  container.innerHTML = html;
}

async function loadAndMount(type) {
  // desmonta módulo atual
  try {
    currentModule?.destroy?.();
  } catch (e) {
    console.error("Erro ao destruir módulo atual:", e);
  }

  clearElement(container);
  currentModule = null;

  if (!type) return;

  try {
    if (type === "PAUTA_MANUAL") {
      const mod = await import("./modules/pautaManual.js");
      currentModule = mod.mount(container);
      return;
    }

    if (type === "RELATORIO_PAUTA_DINAMICA") {
      const mod = await import("./modules/relatorioPautaDinamica.js");
      currentModule = mod.mount(container);
      return;
    }

    showMessage(`<div class="text-danger small">Tipo de documento inválido.</div>`);
  } catch (err) {
    console.error("Falha ao carregar módulo:", err);

    // Mostra o erro na tela (isso vai te dizer exatamente o que está faltando)
    showMessage(`
      <div class="alert alert-danger mb-0">
        <div><strong>Não foi possível carregar o módulo selecionado.</strong></div>
        <div class="small mt-2">
          Verifique se os arquivos existem exatamente nesses caminhos (com mesma maiúscula/minúscula):
          <ul class="mb-2">
            <li><code>./modules/pautaManual.js</code></li>
            <li><code>./modules/relatorioPautaDinamica.js</code></li>
          </ul>
          Erro: <code>${String(err?.message || err)}</code>
        </div>
      </div>
    `);
  }
}

docTypeEl.addEventListener("change", (e) => loadAndMount(e.target.value));
