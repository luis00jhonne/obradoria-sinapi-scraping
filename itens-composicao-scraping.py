import pandas as pd
import requests
import os
import glob
import unicodedata
from datetime import datetime

# ------------------------------
# CONFIGURAÇÕES
# ------------------------------
API_BASE = "http://localhost:8891"
# Produção: API_BASE = "https://api.obradoria.com.br"
ITENS_COMPOSICAO_API_LOTE = f"{API_BASE}/api/itens-composicoes/lote"

# Token JWT da API. Obtenha em POST /api/auth/login e exporte antes de rodar:
#   export OBRADORIA_TOKEN="..."
BEARER_TOKEN = os.getenv("OBRADORIA_TOKEN", "")
AUTH_HEADERS = {
    'Content-Type': 'application/json',
    'Authorization': f'Bearer {BEARER_TOKEN}'
}

if not BEARER_TOKEN:
    print("⚠️  OBRADORIA_TOKEN não definido — a API responderá 401.")

# Diretório das planilhas, relativo a este script (baixar do dataset no Zenodo)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, "excels")

# Arquivos de log, também relativos ao script (não ao diretório de execução)
LOG_FILE = os.path.join(BASE_DIR, "log_envio_itens_composicoes.csv")
ERRO_FILE = os.path.join(BASE_DIR, "erros_envio_itens_composicoes.csv")

# ------------------------------
# FUNÇÕES AUXILIARES
# ------------------------------

def remover_acentos(texto):
    return ''.join(c for c in unicodedata.normalize('NFD', texto) 
                   if unicodedata.category(c) != 'Mn')

def ler_excel(file_path):
    """Lê o Excel e retorna DataFrame."""
    df = pd.read_excel(file_path, sheet_name="Analítico", skiprows=9)
    df.columns = df.columns.str.strip()
    df.columns = df.columns.str.replace('\n', ' ')
    df.columns = [remover_acentos(col) for col in df.columns]
    
    return df

def carregar_log(arquivo):
    """Carrega log de envios."""
    if os.path.exists(arquivo):
        return pd.read_csv(arquivo, dtype={'chave': str})
    return pd.DataFrame(columns=['chave', 'data_envio'])

def salvar_log(log_df, arquivo):
    """Salva log atualizado."""
    log_df.to_csv(arquivo, index=False)

def registrar_erro(chaves, erro):
    """Registra erros em arquivo separado."""
    erro_df = pd.DataFrame({
        'chaves': [','.join(map(str, chaves))],
        'erro': [erro[:500]],
        'data': [datetime.now().strftime('%Y-%m-%d %H:%M:%S')]
    })
    
    if os.path.exists(ERRO_FILE):
        erro_df.to_csv(ERRO_FILE, mode='a', header=False, index=False)
    else:
        erro_df.to_csv(ERRO_FILE, index=False)

# ------------------------------
# PROCESSAMENTO DOS ITENS
# ------------------------------

def processar_itens_composicao(df):
    """Processa DataFrame e extrai itens de composição válidos."""
    itens = []
    chaves = []
    codigo_pai_atual = None
    
    for _, row in df.iterrows():
        # Atualizar código pai quando encontrar uma nova composição principal
        if pd.notna(row.get('Codigo da Composicao')):
            try:
                codigo_pai_atual = int(float(str(row['Codigo da Composicao']).strip()))
            except:
                continue
        
        # Verificar se tem Tipo Item e Código do Item
        tipo_item = row.get('Tipo Item')
        codigo_item = row.get('Codigo do Item')
        
        if pd.isna(tipo_item) or pd.isna(codigo_item):
            continue  # Pula a linha
        
        if not codigo_pai_atual:
            continue  # Não tem código pai definido ainda
        
        tipo_item = str(tipo_item).strip().upper()
        codigo_item = str(codigo_item).strip()

        try:
            codigo_item_int = int(float(str(codigo_item).strip()))
        except:
            continue  # Código inválido, pula
        
        coeficiente = row.get('Coeficiente', 0)
        situacao = row.get('Situacao', 'COM CUSTO')
        descricao = row.get('Descricao')
        unidade_medida = row.get('Unidade')
        
        # Validar coeficiente
        if pd.isna(coeficiente):
            coeficiente = 0
        try:
            coeficiente = float(str(coeficiente).replace(',', '.'))
        except:
            coeficiente = 0
        
        # Criar payload baseado no tipo
        payload = {
            "codigoPai": codigo_pai_atual,
            "codigoFilho": None,
            "codigoInsumo": None,
            "tipoItem": tipo_item,
            "coeficiente": coeficiente,
            "descricao": descricao,
            "unidadeMedida": unidade_medida,
            "situacao": situacao if pd.notna(situacao) else "COM CUSTO"
        }
        
        if tipo_item == "COMPOSICAO":
            payload["codigoFilho"] = codigo_item_int
        elif tipo_item == "INSUMO":
            payload["codigoInsumo"] = codigo_item_int
        else:
            continue  # Tipo inválido, pula
        
        # Criar chave única (codigo_pai|tipo|codigo_item)
        chave = f"{codigo_pai_atual}|{tipo_item}|{codigo_item}"
        
        itens.append(payload)
        chaves.append(chave)
    
    return itens, chaves

# ------------------------------
# ENVIO EM LOTE
# ------------------------------

def enviar_itens_lote(df, log_itens, batch_size=500):
    """Envia itens de composição em lote, pulando já enviados."""
    
    # Processar itens
    itens, chaves = processar_itens_composicao(df)
    
    if not itens:
        print("  ⚠️  Nenhum item válido encontrado no arquivo")
        return log_itens, 0, 0
    
    # Filtrar já enviados
    chaves_enviadas = set(log_itens['chave'].astype(str)) if not log_itens.empty else set()
    
    itens_novos = []
    chaves_novas = []
    
    for item, chave in zip(itens, chaves):
        if chave not in chaves_enviadas:
            itens_novos.append(item)
            chaves_novas.append(chave)
    
    print(f"  📦 Total: {len(itens)} | Já enviados: {len(chaves_enviadas)} | A enviar: {len(itens_novos)}")
    
    if not itens_novos:
        print("  ✓ Todos os itens já foram enviados!")
        return log_itens, len(itens), 0
    
    # Enviar em lotes
    total = len(itens_novos)
    enviados = []
    total_sucesso = 0
    total_falha = 0
    
    for i in range(0, total, batch_size):
        batch = itens_novos[i:i + batch_size]
        batch_chaves = chaves_novas[i:i + batch_size]
        
        try:
            response = requests.post(
                ITENS_COMPOSICAO_API_LOTE,
                json=batch,
                headers=AUTH_HEADERS,
                timeout=60
            )
            
            if response.status_code in [200, 201, 202]:
                enviados.extend(batch_chaves)
                total_sucesso += len(batch)
                print(f"    ✓ Lote {i//batch_size + 1}/{(total + batch_size - 1)//batch_size}: {len(batch)} enviados")
            else:
                total_falha += len(batch)
                registrar_erro(batch_chaves, f"HTTP {response.status_code}: {response.text[:200]}")
                print(f"    ✗ Erro lote {i//batch_size + 1}: HTTP {response.status_code}")
                
        except Exception as e:
            total_falha += len(batch)
            registrar_erro(batch_chaves, str(e))
            print(f"    ✗ Erro conexão lote {i//batch_size + 1}: {e}")
    
    # Atualizar log
    if enviados:
        novos_logs = pd.DataFrame({
            'chave': enviados,
            'data_envio': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        })
        log_itens = pd.concat([log_itens, novos_logs], ignore_index=True)
        salvar_log(log_itens, LOG_FILE)
    
    return log_itens, total_sucesso, total_falha

# ------------------------------
# PROCESSO PRINCIPAL
# ------------------------------

def main():
    print("=" * 70)
    print("INICIANDO PROCESSAMENTO DE ITENS DE COMPOSIÇÃO")
    print("=" * 70)
    
    log_itens = carregar_log(LOG_FILE)
    
    all_files = glob.glob(os.path.join(EXCEL_PATH, "*.xlsx"))
    
    if not all_files:
        print("❌ Nenhum arquivo Excel encontrado em:", EXCEL_PATH)
        return
    
    print(f"\n📁 Encontrados {len(all_files)} arquivo(s)\n")
    
    total_sucesso = 0
    total_falha = 0
    
    for idx, file in enumerate(all_files, 1):
        nome_arquivo = os.path.basename(file)
        print(f"\n[{idx}/{len(all_files)}] 📄 Processando: {nome_arquivo}")
        
        try:
            df = ler_excel(file)
            
            # Debug: mostrar colunas encontradas
            print(f"  🔍 Colunas encontradas: {list(df.columns)[:5]}...")
            
            # Enviar itens
            log_itens, sucessos, falhas = enviar_itens_lote(df, log_itens)
            total_sucesso += sucessos
            total_falha += falhas
            
            print(f"  ✓ Arquivo concluído: {sucessos} sucessos, {falhas} falhas")
            
        except Exception as e:
            print(f"  ✗ Erro ao processar arquivo: {e}")
            import traceback
            traceback.print_exc()
            continue
    
    print("\n" + "=" * 70)
    print("RESUMO FINAL")
    print("=" * 70)
    print(f"✓ Total enviado com sucesso: {total_sucesso}")
    print(f"✗ Total com falha: {total_falha}")
    print(f"📊 Total processado: {total_sucesso + total_falha}")
    
    if total_falha > 0:
        print(f"\n⚠️  Verifique erros em: {ERRO_FILE}")
    
    print("=" * 70)

if __name__ == "__main__":
    main()