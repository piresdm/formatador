import { clearElement } from "./shared/helpers.js";
import { MODULES, MODULES_BY_TYPE } from "./modules/registry.js";

const docTypeEl = document.getElementById("docType");
const container = document.getElementById("moduleContainer");
const extracaoInfoAlertEl = document.getElementById("extracaoInfoAlert");

let currentModule = null;
const EXTRACAO_ALERT_TYPES = new Set([
  "EXTRACAO_PAUTA_PRE_SESSAO",
  "EXTRACAO_PAUTA_POS_SESSAO",
]);

function renderModuleOptions() {
  const placeholder = `<option value="" selected>Selecione...</option>`;
  const options = MODULES.map(
    ({ type, label }) => `<option value="${type}">${label}</option>`,
  );

  docTypeEl.innerHTML = [placeholder, ...options].join("\n");
}

function toggleExtracaoInfoAlert(type) {
  if (!extracaoInfoAlertEl) return;

  const shouldShowAlert = !type || EXTRACAO_ALERT_TYPES.has(type);
  extracaoInfoAlertEl.classList.toggle("d-none", !shouldShowAlert);
}

function showMessage(html) {
  container.innerHTML = html;
}

async function loadAndMount(type) {
  try {
    currentModule?.destroy?.();
  } catch (e) {
    console.error("Erro ao destruir módulo atual:", e);
  }

  clearElement(container);
  currentModule = null;

  if (!type) return;

  const moduleDef = MODULES_BY_TYPE.get(type);

  if (!moduleDef) {
    showMessage('<div class="text-danger small">Tipo de documento inválido.</div>');
    return;
  }

  try {
    const mod = await moduleDef.load();
    currentModule = mod.mount(container);
  } catch (err) {
    console.error("Falha ao carregar módulo:", err);

    const modulePathsList = MODULES.map(({ path }) => `<li><code>${path}</code></li>`).join("");

    showMessage(`
      <div class="alert alert-danger mb-0">
        <div><strong>Não foi possível carregar o módulo selecionado.</strong></div>
        <div class="small mt-2">
          Verifique se os arquivos existem exatamente nesses caminhos (com mesma maiúscula/minúscula):
          <ul class="mb-2">
            ${modulePathsList}
          </ul>
          Erro: <code>${String(err?.message || err)}</code>
        </div>
      </div>
    `);
  }
}

docTypeEl.addEventListener("change", (e) => {
  const { value } = e.target;
  toggleExtracaoInfoAlert(value);
  loadAndMount(value);
});

renderModuleOptions();
docTypeEl.value = "";
toggleExtracaoInfoAlert(docTypeEl.value);
