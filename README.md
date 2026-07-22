# Atualização ARP COMRJ — Atas de Registro de Preços Vigentes

Extrai **todas as Atas de Registro de Preços (ARP) ainda vigentes** do **COMRJ**
(Centro de Obtenção da Marinha no Rio de Janeiro — UASG **771300**), com os
respectivos itens, direto da **API de Dados Abertos do Compras.gov.br**, e gera
uma planilha Excel padronizada.

## Arquivos

| Arquivo | Descrição |
|---|---|
| `extrair_atas_vigentes_comrj.py` | Script de extração (fonte da verdade, reprodutível) |
| `COMRJ_Atas_Vigentes.xlsx` | Planilha gerada (última atualização) |
| `requirements.txt` | Dependências Python |
| `CLAUDE.md` | Contexto para o Claude Code (endpoints, chave de cruzamento, formato) |

## Como atualizar

Basta rodar — a data de referência é **hoje automaticamente** (`HOJE = date.today()`),
e a varredura de anos se ajusta sozinha:

```bash
pip install -r requirements.txt
python extrair_atas_vigentes_comrj.py
```

O arquivo `COMRJ_Atas_Vigentes.xlsx` é sobrescrito com os dados atualizados.
(Para congelar uma data específica, edite `HOJE` no topo do script.)

> **Requisito:** o script faz chamadas HTTP à API `dadosabertos.compras.gov.br`.
> O ambiente onde ele roda precisa ter **acesso de saída à internet** para essa API.

## O que a planilha contém

- **Aba "Atas Vigentes"** — uma linha por item de ata vigente (24 colunas):
  NUP, Nº Pregão, Categoria, Objeto, Situação do Pregão, Modalidade,
  Valor Estimado, Valor Homologado, Nº Item, Descrição do Item, Tipo,
  Situação do Item, Qtd Homolog., Qtd Empenh., Saldo, Fornecedor, CNPJ,
  Valor Unit., Valor Total, Nº ARP, Vig. Início, Vig. Fim, Link PNCP, Extraído em.
- **Aba "Resumo por Pregão"** — uma linha por pregão.
- **Aba "Legenda"** — critérios e conceitos.

## Critério de "vigente"

Mantém apenas itens com `Vig. Fim >= data de referência` e não excluídos. Atas com
vigência encerrada são descartadas. A varredura por vigência inicial cobre
2024–2026, o que captura todas as atas vigentes hoje (uma ARP vai no máximo a
2 anos, então nenhuma ata iniciada antes de meados de 2024 continua vigente).

## Última atualização

- **Data de referência:** 22/07/2026
- **Pregões com atas vigentes:** 77
- **Itens vigentes:** 4.297
