# Obradoria SINAPI Scraping

Sistema de extração e envio automatizado de dados do SINAPI (Sistema Nacional de Pesquisa de Custos e Índices da Construção Civil) para API REST.

## 📚 Contexto Acadêmico

Este projeto foi desenvolvido como parte de uma **dissertação de mestrado** e utilizado na **coleta de dados** para pesquisa sobre custos e orçamentos na construção civil. O sistema automatiza a extração de dados das planilhas SINAPI disponibilizadas pelo Caixa, processando informações de insumos, composições e seus respectivos preços.

---

## 📂 Dataset

As planilhas SINAPI utilizadas neste projeto estão disponíveis como dataset público no Zenodo:

> Luis Jhonne Carvalhal de Melo, & JACOB JUNIOR, A. F. L. (2026). SINAPI Dataset [Data set]. Zenodo. https://doi.org/10.5281/zenodo.18344137

Para utilizar os scripts, baixe o dataset e coloque os arquivos `.xlsx` no diretório `excels/`.

---

## 📋 Índice

- [Dataset](#-dataset)
- [Entidades do Sistema](#-entidades-do-sistema)
- [Estrutura do Projeto](#-estrutura-do-projeto)
- [Requisitos](#-requisitos)
- [Instalação](#-instalação)
- [Configuração](#-configuração)
- [Como Usar](#-como-usar)
- [Sistema de Logs](#-sistema-de-logs)
- [Estrutura dos Dados](#-estrutura-dos-dados)
- [Tratamento de Erros](#-tratamento-de-erros)

---

## 🏗️ Entidades do Sistema

### 1. **Insumos**
Materiais, equipamentos e mão de obra utilizados na construção civil.

**Atributos:**
- `codigo`: Código único do insumo no SINAPI
- `classificacao`: Categoria do insumo (Material, Equipamento, Mão de obra, etc.)
- `nome`: Descrição detalhada do insumo
- `unidadeMedida`: Unidade de medida (kg, m³, h, etc.)

**Fonte:** Planilha **ISD** (Insumos Sem Desoneração)

**Exemplo:**
```json
{
  "codigo": "88261",
  "classificacao": "MATERIAL",
  "nome": "CIMENTO PORTLAND COMPOSTO CP II-32",
  "unidadeMedida": "KG"
}
```

---

### 2. **Composições**
Serviços/atividades da construção civil compostos por múltiplos insumos.

**Atributos:**
- `codigo`: Código único da composição no SINAPI
- `grupo`: Grupo/categoria da composição
- `nome`: Descrição do serviço
- `unidadeMedida`: Unidade de medida do serviço (m², m³, un, etc.)

**Fonte:** Planilha **CSD** (Composições Sem Desoneração)

**Exemplo:**
```json
{
  "codigo": "94965",
  "grupo": "ALVENARIA",
  "nome": "ALVENARIA DE VEDACAO DE BLOCOS CERAMICOS FURADOS",
  "unidadeMedida": "M2"
}
```

---

### 3. **Preços de Insumos**
Valores dos insumos por estado e mês de referência, com e sem desoneração.

**Atributos:**
- `codigoInsumo`: Código do insumo
- `origemPreco`: Origem do preço (CR - Cotação Regional, etc.)
- `mesAnoReferencia`: Período de referência (ex: "09/2025")
- `dataEmissao`: Data de emissão da tabela
- `estado`: Sigla do estado (UF)
- `coeficienteSemDesoneracao`: Preço sem desoneração
- `coeficienteComDesoneracao`: Preço com desoneração (opcional)

**Fontes:**
- **ISD** (Insumos Sem Desoneração)
- **ICD** (Insumos Com Desoneração)

**Exemplo:**
```json
{
  "codigoInsumo": 88261,
  "origemPreco": "CR",
  "mesAnoReferencia": "11/2025",
  "dataEmissao": "2025-12-09",
  "estado": "SP",
  "coeficienteSemDesoneracao": 0.85,
  "coeficienteComDesoneracao": 0.82
}
```

---

### 4. **Preços de Composições**
Custos das composições por estado e mês de referência, incluindo BDI (Benefícios e Despesas Indiretas).

**Atributos:**
- `codigoComposicao`: Código da composição
- `mesAnoReferencia`: Período de referência
- `dataEmissao`: Data de emissão da tabela
- `estado`: Sigla do estado (UF)
- `custoSemDesoneracao`: Custo direto sem desoneração
- `acSemDesoneracao`: BDI (AC%) sem desoneração
- `custoComDesoneracao`: Custo direto com desoneração (opcional)
- `acComDesoneracao`: BDI (AC%) com desoneração (opcional)

**Fontes:**
- **CSD** (Composições Sem Desoneração)
- **CCD** (Composições Com Desoneração)

**Exemplo:**
```json
{
  "codigoComposicao": 94965,
  "mesAnoReferencia": "11/2025",
  "dataEmissao": "2025-12-09",
  "estado": "RJ",
  "custoSemDesoneracao": 125.50,
  "acSemDesoneracao": 28.30,
  "custoComDesoneracao": 120.00,
  "acComDesoneracao": 27.50
}
```

---

### 5. **Itens de Composição**
Relacionamento entre composições e insumos, definindo quantidades utilizadas.

**Atributos:**
- `codigoComposicao`: Código da composição
- `tipo`: Tipo do item (I - Insumo, C - Composição)
- `codigoItem`: Código do insumo ou composição auxiliar
- `coeficiente`: Quantidade do insumo necessária

**Fonte:** Planilha **Analítico**

**Exemplo:**
```json
{
  "codigoComposicao": 94965,
  "tipo": "I",
  "codigoItem": 88261,
  "coeficiente": 5.5
}
```

---

## 📁 Estrutura do Projeto

```
obradoria-sinapi-scraping/
├── excels/                          # Planilhas SINAPI (baixar do Zenodo)
├── insumo-scraping.py               # Script para extrair insumos
├── composicao-scraping.py           # Script para extrair composições
├── preco-insumo-scraping.py         # Script para extrair preços de insumos
├── preco-composicao-scraping.py     # Script para extrair preços de composições
├── itens-composicao-scraping.py     # Script para extrair itens das composições
├── log_envio_insumos.csv            # Log de insumos enviados
├── log_envio_composicoes.csv        # Log de composições enviadas
├── log_envio_precos_insumos.csv     # Log de preços de insumos enviados
├── log_envio_precos_composicoes.csv # Log de preços de composições enviados
├── log_envio_itens.csv              # Log de itens enviados
├── erros_envio*.csv                 # Logs de erros (gerados automaticamente)
└── README.md                        # Este arquivo
```

---

## 🔧 Requisitos

### Software
- Python 3.8+
- API REST do `obradoria-backend` acessível (porta 8891 em execução local)
- Token JWT válido da API (ver [Autenticação](#autenticação))

### Bibliotecas Python
```bash
pandas
requests
openpyxl
```

---

## 📦 Instalação

1. **Clone ou baixe o projeto:**
```bash
cd /caminho/para/projeto
```

2. **Instale as dependências:**
```bash
pip install pandas requests openpyxl
```

3. **Baixe o dataset SINAPI do Zenodo:**
```bash
mkdir -p excels
```
   - Acesse: https://doi.org/10.5281/zenodo.18344137
   - Baixe os arquivos `.xlsx` do dataset
   - Coloque os arquivos no diretório `excels/`

---

## ⚙️ Configuração

### Endpoints da API

Edite os scripts para configurar os endpoints da sua API:

Todos os scripts apontam para **`http://localhost:8891`**, definido na constante
`API_BASE` no topo do arquivo. Para enviar a outro ambiente, troque essa linha —
a URL de produção está logo abaixo, comentada:

```python
API_BASE = "http://localhost:8891"
# Produção: API_BASE = "https://api.obradoria.com.br"
```

Cada script monta seu endpoint a partir dela:

| Script | Endpoint |
| ------ | -------- |
| `insumo-scraping.py` | `{API_BASE}/api/insumos/lote` |
| `composicao-scraping.py` | `{API_BASE}/api/composicoes/lote` |
| `preco-insumo-scraping.py` | `{API_BASE}/api/preco-insumos/lote` |
| `preco-composicao-scraping.py` | `{API_BASE}/api/preco-composicoes/lote` |
| `itens-composicao-scraping.py` | `{API_BASE}/api/itens-composicoes/lote` |

### Autenticação

A API exige **JWT** em todos os endpoints (exceto `/api/auth/login`). Os cinco
scripts leem o token da variável de ambiente `OBRADORIA_TOKEN`:

```bash
export OBRADORIA_TOKEN=$(curl -s -X POST http://localhost:8891/api/auth/login \
  -H "Content-Type: application/json" \
  -d '{"email":"seu@email","senha":"..."}' | jq -r .token)
```

Sem a variável definida, o script avisa na saída e a API responde **401**.

⚠️ **Nunca coloque o token no código** — este repositório é público.

### Caminho dos Arquivos Excel

Os scripts procuram as planilhas em `excels/`, **relativo ao próprio script**:

```python
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, "excels")
```

Os arquivos de log e de erro seguem a mesma regra e são gravados na pasta do
repositório. Não é preciso ajustar nada, e os scripts podem ser executados de
qualquer diretório.

---

## 🚀 Como Usar

### Ordem Recomendada de Execução

**⚠️ IMPORTANTE:** Execute os scripts na ordem abaixo para manter a integridade referencial:

#### 1. Extrair Insumos
```bash
python insumo-scraping.py
```
**Saída esperada:**
```
======================================================================
INICIANDO PROCESSAMENTO DE INSUMOS
======================================================================

📁 Encontrados 12 arquivo(s)

[1/12] 📄 Processando: SINAPI_Referência_2024_09.xlsx
  📦 Total: 8542 | Já enviados: 0 | A enviar: 8542
    ✓ Lote 1/18: 500 enviados
    ✓ Lote 2/18: 500 enviados
    ...
  ✓ Arquivo concluído: 8542 sucessos, 0 falhas

======================================================================
RESUMO FINAL
======================================================================
✓ Total enviado com sucesso: 102504
✗ Total com falha: 0
📊 Total processado: 102504
======================================================================
```

---

#### 2. Extrair Composições
```bash
python composicao-scraping.py
```
**Saída esperada:**
```
============================================================
INICIANDO PROCESSAMENTO DE COMPOSIÇÕES
============================================================

📁 Encontrados 12 arquivo(s)

[1/12] 📄 Processando: SINAPI_Referência_2024_09.xlsx
  📊 Total: 6234 | Já enviados: 0 | A enviar: 6234
  ✓ Lote 1/13: 500 enviados
  ✓ Lote 2/13: 500 enviados
  ...
  ✓ Concluído: 6234 sucessos, 0 falhas

============================================================
RESUMO FINAL
============================================================
✓ Total enviado com sucesso: 74808
✗ Total com falha: 0
📊 Total processado: 74808
============================================================
```

---

#### 3. Extrair Preços de Insumos
```bash
python preco-insumo-scraping.py
```
**Saída esperada:**
```
======================================================================
INICIANDO PROCESSAMENTO DE PREÇOS DE INSUMOS
======================================================================

📁 Encontrados 12 arquivo(s)

[1/12] 📄 Processando: SINAPI_Referência_2024_09.xlsx
  📅 Mês: 09/2024 | Data: 2024-10-09
  📊 ISD: 8542 linhas | UFs: AC, AL, AM, AP, BA, CE, DF, ES, ...
  📊 ICD: 8542 linhas
  💰 Total: 230634 | Já enviados: 0 | A enviar: 230634
    ✓ Lote 1/462: 500 enviados
    ✓ Lote 2/462: 500 enviados
    ...
  ✓ Arquivo concluído: 230634 sucessos, 0 falhas

======================================================================
RESUMO FINAL
======================================================================
✓ Total enviado com sucesso: 2767608
✗ Total com falha: 0
📊 Total processado: 2767608
======================================================================
```

---

#### 4. Extrair Preços de Composições
```bash
python preco-composicao-scraping.py
```
**Saída esperada:**
```
======================================================================
INICIANDO PROCESSAMENTO DE PREÇOS DE COMPOSIÇÕES
======================================================================

📁 Encontrados 12 arquivo(s)

[1/12] 📄 Processando: SINAPI_Referência_2024_09.xlsx
  📅 Mês: 09/2024 | Data: 2024-10-09
  📊 CSD: 6234 linhas | UFs: AC, AL, AM, AP, BA, CE, DF, ES, ...
  📊 CCD: 6234 linhas
  💰 Total: 168552 | Já enviados: 0 | A enviar: 168552
    ✓ Lote 1/338: 500 enviados
    ✓ Lote 2/338: 500 enviados
    ...
  ✓ Arquivo concluído: 168552 sucessos, 0 falhas

======================================================================
RESUMO FINAL
======================================================================
✓ Total enviado com sucesso: 2022624
✗ Total com falha: 0
📊 Total processado: 2022624
======================================================================
```

---

#### 5. Extrair Itens de Composição
```bash
python itens-composicao-scraping.py
```
**Saída esperada:**
```
======================================================================
INICIANDO PROCESSAMENTO DE ITENS DE COMPOSIÇÃO
======================================================================

📁 Encontrados 12 arquivo(s)

[1/12] 📄 Processando: SINAPI_Referência_2024_09.xlsx
  📊 Total de linhas: 52341
  📊 Composições únicas: 6234
  📊 Itens extraídos: 45678
  💰 Total: 45678 | Já enviados: 0 | A enviar: 45678
    ✓ Lote 1/92: 500 enviados
    ✓ Lote 2/92: 500 enviados
    ...
  ✓ Arquivo concluído: 45678 sucessos, 0 falhas

======================================================================
RESUMO FINAL
======================================================================
✓ Total enviado com sucesso: 548136
✗ Total com falha: 0
📊 Total processado: 548136
======================================================================
```

---

### Executar Todos os Scripts Sequencialmente

```bash
python insumo-scraping.py && \
python composicao-scraping.py && \
python preco-insumo-scraping.py && \
python preco-composicao-scraping.py && \
python itens-composicao-scraping.py
```

---

## 📝 Sistema de Logs

### Como Funciona

Cada script mantém um **arquivo de log CSV** para registrar os dados já enviados. Isso evita:
- ✅ Duplicação de dados na API
- ✅ Reenvio desnecessário em caso de interrupção
- ✅ Processamento incremental de novos arquivos

### Arquivos de Log

| Script | Arquivo de Log | Chave Única |
|--------|---------------|-------------|
| `insumo-scraping.py` | `log_envio_insumos.csv` | `codigo` |
| `composicao-scraping.py` | `log_envio_composicoes.csv` | `codigo` |
| `preco-insumo-scraping.py` | `log_envio_precos_insumos.csv` | `codigo\|uf\|mes_ref` |
| `preco-composicao-scraping.py` | `log_envio_precos_composicoes.csv` | `codigo\|uf\|mes_ref` |
| `itens-composicao-scraping.py` | `log_envio_itens.csv` | `cod_comp\|tipo\|cod_item` |

### Estrutura dos Logs

#### log_envio_insumos.csv
```csv
codigo,data_envio
88261,2025-01-10 14:30:25
88262,2025-01-10 14:30:25
```

#### log_envio_precos_insumos.csv
```csv
chave,mes_referencia,data_envio
88261|SP|09/2024,09/2024,2025-01-10 14:35:10
88261|RJ|09/2024,09/2024,2025-01-10 14:35:10
```

### Limpar Logs (Reprocessar Tudo)

Para reenviar todos os dados, **delete os arquivos de log**:

```bash
rm log_envio_*.csv
```

⚠️ **Atenção:** Isso causará duplicação de dados se a API não tiver validação de chave única!

---

## 📊 Estrutura dos Dados

### Planilhas SINAPI

As planilhas Excel do SINAPI possuem as seguintes abas:

| Aba | Conteúdo | Script |
|-----|----------|--------|
| **ISD** | Insumos Sem Desoneração | `insumo-scraping.py`<br>`preco-insumo-scraping.py` |
| **ICD** | Insumos Com Desoneração | `preco-insumo-scraping.py` |
| **CSD** | Composições Sem Desoneração | `composicao-scraping.py`<br>`preco-composicao-scraping.py` |
| **CCD** | Composições Com Desoneração | `preco-composicao-scraping.py` |
| **Analítico** | Itens de Composição | `itens-composicao-scraping.py` |

### Particularidades da Extração

#### 1. **Fórmulas HYPERLINK**
As colunas "Código da Composição" contêm fórmulas HYPERLINK em vez de valores diretos:

```excel
=HYPERLINK("#"&CELL("address",OFFSET(Analítico!$B$1,MATCH(105003,Analítico!$B:$B,0)-1,3)),105003)
```

Os scripts detectam e extraem automaticamente o código (105003) destas fórmulas usando regex.

#### 2. **Múltiplos UFs por Linha**
As planilhas de preços possuem uma coluna para cada estado (27 UFs). Os scripts:
- Detectam automaticamente as colunas de UF
- Transformam cada linha em múltiplos registros (um por UF)
- Enviam apenas preços com valores válidos (não-nulos e não-zero)

#### 3. **Desoneração Opcional**
Nem todas as planilhas possuem dados COM desoneração (ICD/CCD). Os scripts:
- Tentam ler ambas as abas
- Processam apenas a aba SEM desoneração se COM desoneração não existir
- Enviam `null` para campos de desoneração quando não disponíveis

---

## ❌ Tratamento de Erros

### Logs de Erro

Erros durante o envio são registrados em arquivos separados:

```
erros_envio_insumos.csv
erros_envio_composicoes.csv
erros_envio_precos_insumos.csv
erros_envio_precos_composicoes.csv
erros_envio_itens.csv
```

### Estrutura dos Arquivos de Erro

```csv
codigos,erro,data
"88261,88262,88263",HTTP 500: Internal Server Error,2025-01-10 14:35:45
```

### Tipos de Erro Comuns

| Erro | Causa | Solução |
|------|-------|---------|
| `Connection refused` | API não está rodando | Inicie a API na porta 8891 |
| `HTTP 400` | Dados inválidos/incompletos | Verifique formato dos dados |
| `HTTP 500` | Erro interno da API | Verifique logs da API |
| `Timeout` | Lote muito grande ou API lenta | Reduza `batch_size` |

### Reprocessar Erros

Os itens com erro **NÃO são registrados no log**. Para reprocessá-los:

1. Corrija a causa do erro
2. Execute o script novamente
3. Os itens com erro serão reprocessados automaticamente

---

## 🔍 Validação dos Dados

### Verificar Totais Enviados

```bash
wc -l log_envio_*.csv
```

### Verificar Erros

```bash
cat erros_envio_*.csv
```
---

## 🛠️ Troubleshooting

### Problema: "Nenhum arquivo Excel encontrado"

**Solução:**
1. Verifique se os arquivos estão no diretório `excels/`
2. Confirme que são arquivos `.xlsx` (não `.xls`)
3. Ajuste `EXCEL_PATH` nos scripts se necessário

---

### Problema: "Código da Composição está None"

**Solução:**
- Isso ocorre quando a fórmula HYPERLINK não pode ser extraída
- Verifique se o arquivo Excel está corrompido
- Reexporte o arquivo Excel original

---

### Problema: "Poucos preços enviados"

**Causa:** Preços zerados ou nulos são ignorados automaticamente.

**Solução:** Isso é comportamento esperado. Nem todos os insumos/composições têm preços em todos os estados.

---

### Problema: "Lentidão ao processar"

**Solução:**
1. Reduza `batch_size` de 500 para 100-200
2. Aumente timeout da API
3. Processe arquivos individualmente

---

## 📄 Licença e Uso Acadêmico

Este código foi desenvolvido para fins de pesquisa acadêmica. Os dados do SINAPI são de domínio público e disponibilizados pela Caixa https://www.caixa.gov.br/poder-publico/modernizacao-gestao/sinapi/Paginas/default.aspx.

**Citação do Dataset:**
```
Luis Jhonne Carvalhal de Melo, & JACOB JUNIOR, A. F. L. (2026).
SINAPI Dataset [Data set]. Zenodo. https://doi.org/10.5281/zenodo.18344137
```

---

## 🤝 Contribuições

Para melhorias ou correções:
1. Documente o problema encontrado
2. Teste a solução localmente
3. Atualize este README se necessário

---

## 📞 Suporte

Para dúvidas sobre:
- **Dados SINAPI:** [https://www.caixa.gov.br/poder-publico/modernizacao-gestao/sinapi/Paginas/default.aspx](https://www.caixa.gov.br/poder-publico/modernizacao-gestao/sinapi/Paginas/default.aspx)
- **API REST:** Consulte documentação do backend
- **Scripts Python:** Verifique comentários inline no código

---

**Desenvolvido com 🏗️ para pesquisa em Engenharia de Computação**

*Última atualização: Janeiro 2026*
