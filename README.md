# Formatador de Documentos DJULG

Aplicação **100% front-end** para formatação de documentos e extração de planilhas.

## Importante
- O projeto **não usa backend**.
- Não precisa de API, banco ou servidor de aplicação.
- A geração de `.xlsx`, `.docx` e `.pdf` acontece no navegador.

## Módulos
- Relatório Pauta Dinâmica
- Pauta Manual
- Extração Pauta (pré-sessão)
- Extração Pauta (pós-sessão)

## Como usar
1. Abra `index.html` no navegador (ou via qualquer servidor estático de sua preferência).
2. Selecione o módulo desejado no campo **Tipo de documento**.
3. Envie os arquivos solicitados pelo módulo.
4. Clique no botão de geração do arquivo.

## Observação sobre execução local
Em alguns navegadores, recursos ES Modules podem ter restrições no `file://`.
Se isso acontecer, basta abrir o projeto com **qualquer servidor estático** (não é backend da aplicação, apenas hospedagem de arquivos estáticos).
