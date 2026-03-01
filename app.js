import { clearElement } from "./shared/helpers.js";

const docTypeEl = document.getElementById("docType");
const container = document.getElementById("moduleContainer");

let currentModule = null;

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

  // fallback
  container.innerHTML = `<div class="text-danger small">Tipo de documento inválido.</div>`;
}

docTypeEl.addEventListener("change", (e) => {
  loadAndMount(e.target.value);
});
