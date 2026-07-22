# Contexto do projeto — Atualização ARP COMRJ

Extração de Atas de Registro de Preços (ARP) **vigentes** do **COMRJ**
(UASG **771300**) via API de Dados Abertos do Compras.gov.br, gerando
`COMRJ_Atas_Vigentes.xlsx`. Script principal: `extrair_atas_vigentes_comrj.py`.

## Como atualizar até a data de hoje
Edite `HOJE = date(AAAA, M, D)` no topo do script para a data da extração e rode
`python extrair_atas_vigentes_comrj.py`. A planilha é sobrescrita.

## API (base `https://dadosabertos.compras.gov.br`)
`tamanhoPagina` deve ficar entre 10 e 500 (fora disso = HTTP 400).

- **Itens de ARP:** `GET /modulo-arp/2_consultarARPItem?codigoUnidadeGerenciadora=771300&dataVigenciaInicialMin=AAAA-01-01&dataVigenciaInicialMax=AAAA-12-31&pagina=N&tamanhoPagina=500`. Só filtra por `dataVigenciaInicial`; varrer anos 2024–2026 cobre todas as atas vigentes hoje.
- **Contratações (objeto/NUP/situação/valores):** `GET /modulo-contratacoes/1_consultarContratacoes_PNCP_14133?unidadeOrgaoCodigoUnidade=771300&dataPublicacaoPncpInicial=...&dataPublicacaoPncpFinal=...&codigoModalidade=M`. Carregar modalidades 5, 6 e 7.

## Regras não óbvias (não quebrar)
- **Cruzamento pela CHAVE CANÔNICA do PNCP:** `item.numeroControlePncpCompra == contratacao.numeroControlePNCP`. Cruzar por (numeroCompra, ano) falha muito porque `anoCompra` (item) ≠ `anoCompraPncp` (contratação).
- **Vigente:** `dataVigenciaFinal >= HOJE` e `itemExcluido == False`.
- **Uma linha por item (sem cadastro reserva):** dedup por `(numeroControlePncpAta, numeroItem)` mantendo menor `classificacaoFornecedor` (001 = vencedor).
- **Link da ata:** parsear `numeroControlePncpAta` (`cnpj-1-seqCompra/ano-seqAta`) → `https://pncp.gov.br/app/atas/{cnpj}/{ano}/{seqCompra}/{seqAta}`.
- **Categoria** não vem da API: lookup manual por (numeroCompra, ano) + heurística por palavra-chave do objeto. Ao encontrar pregões novos sem categoria, adicionar ao `CAT_LOOKUP`/`CAT_KEYWORDS`.
- A API de consulta do PNCP (`pncp.gov.br/api/consulta/v1`) é lenta/instável — usar só como fallback e falhar rápido.

## Formato da planilha (24 colunas, nesta ordem)
NUP · Nº Pregão · Categoria · Objeto · Situação do Pregão · Modalidade ·
Valor Estimado (R$) · Valor Homologado (R$) · Nº Item · Descrição do Item · Tipo ·
Situação do Item · Qtd Homolog. · Qtd Empenh. · Saldo · Fornecedor · CNPJ ·
Valor Unit. (R$) · Valor Total (R$) · Nº ARP · Vig. Início · Vig. Fim · Link PNCP · Extraído em.

Abas: "Atas Vigentes", "Resumo por Pregão", "Legenda". Header azul-marinho (003366),
valores monetários em `R$ #,##0.00`, `freeze_panes` A2.
