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

## Critério de "vigente" e PRORROGAÇÕES

Mantém apenas itens com `Vig. Fim >= data de referência`, item não excluído e ata
não cancelada.

> **Importante — atas prorrogadas/renovadas:** a API de dados abertos guarda a
> vigência **original** e NÃO reflete prorrogações. Por isso o script confirma a
> vigência de cada ata no **PNCP** (`/api/pncp/v1/.../atas`), que é a fonte
> autoritativa, e usa essa data. Atas renovadas recebem a situação
> **"Vigente (prorrogada)"** (destacadas em azul). Sem esse cruzamento, atas
> renovadas seriam perdidas assim que a data original vencesse.

## Última atualização

- **Data de referência:** 21/09/2026
- **Pregões com atas vigentes:** 71
- **Itens vigentes:** 3.708 (dos quais **195 em atas prorrogadas**, 25 atas)
- **Valor total dos itens vigentes:** R$ 1.247.324.464,17
