# CLAUDE.md — obradoria-sinapi-scraping

Orientações para o Claude Code neste repositório.
Contexto do produto e regras gerais: `../CLAUDE.md` do workspace.

## Visão geral

Scripts de carga que leem as planilhas SINAPI (Caixa) e alimentam a API do
ObradorIA pelos endpoints `/lote`. É processo **offline e incremental**: cada
script mantém um CSV de log e nunca reenvia o que já foi enviado.

Origem dos dados: dataset no Zenodo — DOI `10.5281/zenodo.18344137`.
Os `.xlsx` vão em `excels/` e **não são versionados**.

## Comandos

```bash
pip install pandas requests openpyxl

# token JWT obrigatório — os scripts leem de OBRADORIA_TOKEN
export OBRADORIA_TOKEN=$(curl -s -X POST http://localhost:8891/api/auth/login \
  -H "Content-Type: application/json" \
  -d '{"email":"...","senha":"..."}' | jq -r .token)

# ordem obrigatória — respeita a integridade referencial
python insumo-scraping.py
python composicao-scraping.py
python preco-insumo-scraping.py
python preco-composicao-scraping.py
python itens-composicao-scraping.py
```

O backend precisa estar rodando e alcançável. Os cinco scripts usam
`API_BASE = "http://localhost:8891"`, com a URL de produção comentada logo abaixo.

## Scripts e correspondências

| Script | Aba da planilha | Endpoint | Log |
| ------ | --------------- | -------- | --- |
| `insumo-scraping.py` | ISD | `/api/insumos/lote` | `log_envio_insumos.csv` |
| `composicao-scraping.py` | CSD | `/api/composicoes/lote` | `log_envio_composicoes.csv` |
| `preco-insumo-scraping.py` | ISD + ICD | `/api/preco-insumos/lote` | `log_envio_precos_insumos.csv` |
| `preco-composicao-scraping.py` | CSD + CCD | `/api/preco-composicoes/lote` | `log_envio_precos_composicoes.csv` |
| `itens-composicao-scraping.py` | Analítico | `/api/itens-composicao/lote` | `log_envio_itens.csv` |

ISD/CSD = sem desoneração; ICD/CCD = com desoneração (nem toda planilha tem).

## Convenções dos scripts

Todos seguem a mesma forma; ao editar um, mantenha o padrão dos demais:

- Constantes de configuração no topo (`API_BASE`, `EXCEL_PATH`, `*_API_LOTE`, `batch_size=500`).
- Token JWT sempre de `os.getenv("OBRADORIA_TOKEN")` — **nunca literal no código**,
  o repositório é público. Não imprima `AUTH_HEADERS` em log.
- Toda requisição envia `headers=AUTH_HEADERS`.
- Lote de 500 registros por requisição.
- Log CSV consultado **antes** do envio, gravado **depois** do sucesso.
- Falhas vão para `erros_envio_*.csv` e **não** entram no log — reexecutar reprocessa.
- Saída no terminal com progresso por arquivo e resumo final.

## Armadilhas

- **Todos os caminhos são relativos ao próprio script**, via
  `BASE_DIR = os.path.dirname(os.path.abspath(__file__))`: `EXCEL_PATH`, os
  `log_envio_*.csv` e os `erros_envio_*.csv`. Os scripts funcionam de qualquer
  diretório de execução — não reintroduza caminho literal nem absoluto.
- **Fórmulas HYPERLINK**: a coluna de código da composição traz fórmula, não
  valor. Os scripts extraem o número por regex — preserve isso.
- **Uma coluna por UF** nas planilhas de preço: cada linha vira até 27 registros.
  Preços nulos ou zerados são descartados de propósito.
- **Apagar os `log_envio_*.csv` reenvia tudo** e duplica dados se a API não tiver
  chave única. Nunca faça isso sem confirmar com o usuário.
- **Volume alto**: preços de insumos passam de 2,7 milhões de registros. Mudança
  no `batch_size` ou no timeout tem efeito grande no tempo total.

## Contexto acadêmico

O repositório faz parte de uma dissertação de mestrado. O `README.md` documenta o
formato dos dados e traz a citação do dataset — mantenha-o atualizado ao mudar as
entidades enviadas.
