const createModuleDef = ({ type, label, showExtracaoInfo, path }) => ({
  type,
  label,
  showExtracaoInfo,
  path,
  load: () => import(path),
});

export const MODULES = [
  createModuleDef({
    type: "LINKS",
    label: "Links",
    showExtracaoInfo: false,
    path: "./links.js",
  }),
  createModuleDef({
    type: "RELATORIO_PAUTA_DINAMICA",
    label: "Relatório Pauta Dinâmica",
    showExtracaoInfo: false,
    path: "./relatorioPautaDinamica.js",
  }),
  createModuleDef({
    type: "PAUTA_MANUAL",
    label: "Pauta Manual",
    showExtracaoInfo: false,
    path: "./pautaManual.js",
  }),
  createModuleDef({
    type: "EXTRACAO_PAUTA_PRE_SESSAO",
    label: "Extração Pauta (pré-sessão)",
    showExtracaoInfo: true,
    path: "./extracaoPautaPreSessao.js",
  }),
  createModuleDef({
    type: "EXTRACAO_PAUTA_POS_SESSAO",
    label: "Extração Pauta (pós-sessão)",
    showExtracaoInfo: true,
    path: "./extracaoPautaPosSessao.js",
  }),
];

export const MODULES_BY_TYPE = new Map(MODULES.map((moduleDef) => [moduleDef.type, moduleDef]));
