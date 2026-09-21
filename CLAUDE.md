# Contexto do projeto — Atualização ARP COMRJ

Extração de Atas de Registro de Preços (ARP) **vigentes** do **COMRJ**
(UASG **771300**) via API de Dados Abertos do Compras.gov.br, gerando
`COMRJ_Atas_Vigentes.xlsx`. Script principal: `extrair_atas_vigentes_comrj.py`.

## Como atualizar até a data de hoje
Só rodar `python extrair_atas_vigentes_comrj.py` — a data de referência já é
`date.today()` e a varredura de anos se ajusta sozinha. A planilha é sobrescrita.
Depois, commit + push. (Para congelar uma data, edite `HOJE` no topo do script.)

**Atenção:** exige acesso de saída à internet para `dadosabertos.compras.gov.br`.
Se o sandbox bloquear a saída, o script falha na coleta.

## API (base `https://dadosabertos.compras.gov.br`)
`tamanhoPagina` deve ficar entre 10 e 500 (fora disso = HTTP 400).

- **Itens de ARP:** `GET /modulo-arp/2_consultarARPItem?codigoUnidadeGerenciadora=771300&dataVigenciaInicialMin=AAAA-01-01&dataVigenciaInicialMax=AAAA-12-31&pagina=N&tamanhoPagina=500`. Só filtra por `dataVigenciaInicial`; varrer anos 2024–2026 cobre todas as atas vigentes hoje.
- **Contratações (objeto/NUP/situação/valores):** `GET /modulo-contratacoes/1_consultarContratacoes_PNCP_14133?unidadeOrgaoCodigoUnidade=771300&dataPublicacaoPncpInicial=...&dataPublicacaoPncpFinal=...&codigoModalidade=M`. Carregar modalidades 5, 6 e 7.

## Regras não óbvias (não quebrar)
- **PRORROGAÇÕES / vigência autoritativa:** a API de dados abertos (`modulo-arp`) NÃO reflete prorrogações — guarda a vigência ORIGINAL. Para cada compra, consultar o PNCP `GET https://pncp.gov.br/api/pncp/v1/orgaos/{cnpj}/compras/{ano}/{seq}/atas` (campos `dataVigenciaFim`, `dataVigenciaInicio`, `cancelado`; chave `numeroControlePNCP` == `numeroControlePncpAta` dos dados abertos) e usar essa vigência. Ex.: ata 00154/2026 aparece como 13/10/2026 nos dados abertos, mas 13/10/2027 no PNCP. O endpoint em lote `/api/consulta/v1/atas` costuma dar timeout — usar a consulta por compra. Item prorrogado recebe situação "Vigente (prorrogada)".
- **Cruzamento pela CHAVE CANÔNICA do PNCP:** `item.numeroControlePncpCompra == contratacao.numeroControlePNCP`. Cruzar por (numeroCompra, ano) falha muito porque `anoCompra` (item) ≠ `anoCompraPncp` (contratação).
- **Vigente:** `dataVigenciaFinal (PNCP) >= HOJE`, `itemExcluido == False` e ata `cancelado == False`.
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
