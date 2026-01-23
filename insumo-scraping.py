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
INSUMO_API_LOTE = f"{API_BASE}/api/insumos/lote"

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
LOG_INSUMOS_FILE = os.path.join(BASE_DIR, "log_envio_insumos.csv")
ERRO_FILE = os.path.join(BASE_DIR, "erros_envio_insumos.csv")

# ------------------------------
# FUNÇÕES AUXILIARES
# ------------------------------

def remover_acentos(texto):
    return ''.join(c for c in unicodedata.normalize('NFD', texto) 
                   if unicodedata.category(c) != 'Mn')

def ler_excel(file_path):
    """Lê o Excel e retorna DataFrame."""
    df = pd.read_excel(file_path, sheet_name="ISD", skiprows=9)
    df.columns = df.columns.str.strip()
    df.columns = df.columns.str.replace('\n', ' ')
    df.columns = [remover_acentos(col) for col in df.columns]
    
    return df

def carregar_log(arquivo):
    """Carrega log de envios."""
    if os.path.exists(arquivo):
        return pd.read_csv(arquivo, dtype={'codigo': str})
    return pd.DataFrame(columns=['codigo', 'data_envio'])

def salvar_log(log_df, arquivo):
    """Salva log atualizado."""
    log_df.to_csv(arquivo, index=False)

def registrar_erro(codigos, erro):
    """Registra erros em arquivo separado."""
    erro_df = pd.DataFrame({
        'codigos': [','.join(map(str, codigos))],
        'erro': [erro[:500]],  # limita tamanho do erro
        'data': [datetime.now().strftime('%Y-%m-%d %H:%M:%S')]
    })
    
    if os.path.exists(ERRO_FILE):
        erro_df.to_csv(ERRO_FILE, mode='a', header=False, index=False)
    else:
        erro_df.to_csv(ERRO_FILE, index=False)

# ------------------------------
# ENVIO EM LOTE - INSUMOS
# ------------------------------

def enviar_insumos_lote(df, log_insumos, batch_size=500):
    """Envia insumos em lote, pulando já enviados."""
    
    df['codigo_str'] = df['Codigo do Insumo'].astype(str).str.strip()
    
    # Filtrar já enviados
    codigos_enviados = set(log_insumos['codigo'].astype(str)) if not log_insumos.empty else set()
    df_novos = df[~df['codigo_str'].isin(codigos_enviados)]
    
    print(f"  📦 Total: {len(df)} | Já enviados: {len(codigos_enviados)} | A enviar: {len(df_novos)}")
    
    if df_novos.empty:
        print("  ✓ Todos os insumos já foram enviados!")
        return log_insumos, len(df), 0
    
    # Preparar payloads
    payloads = []
    codigos = []
    
    for _, row in df_novos.iterrows():
        codigo = row['codigo_str']
        payload = {
            "classificacao": row["Classificacao"],
            "codigo": codigo,
            "nome": row["Descricao do Insumo"],
            "unidadeMedida": row["Unidade"]
        }
        payloads.append(payload)
        codigos.append(codigo)
    
    # Enviar em lotes
    total = len(payloads)
    enviados = []
    total_sucesso = 0
    total_falha = 0
    
    for i in range(0, total, batch_size):
        batch = payloads[i:i + batch_size]
        batch_codigos = codigos[i:i + batch_size]
        
        try:
            response = requests.post(
                INSUMO_API_LOTE,
                json=batch,
                headers=AUTH_HEADERS,
                timeout=60
            )
            
            if response.status_code in [200, 201, 202]:
                enviados.extend(batch_codigos)
                total_sucesso += len(batch)
                print(f"    ✓ Lote {i//batch_size + 1}/{(total + batch_size - 1)//batch_size}: {len(batch)} enviados")
            else:
                total_falha += len(batch)
                registrar_erro(batch_codigos, f"HTTP {response.status_code}: {response.text[:200]}")
                print(f"    ✗ Erro lote {i//batch_size + 1}: HTTP {response.status_code}")
                
        except Exception as e:
            total_falha += len(batch)
            registrar_erro(batch_codigos, str(e))
            print(f"    ✗ Erro conexão lote {i//batch_size + 1}: {e}")
    
    # Atualizar log
    if enviados:
        novos_logs = pd.DataFrame({
            'codigo': enviados,
            'data_envio': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        })
        log_insumos = pd.concat([log_insumos, novos_logs], ignore_index=True)
        salvar_log(log_insumos, LOG_INSUMOS_FILE)
    
    return log_insumos, total_sucesso, total_falha

# ------------------------------
# PROCESSO PRINCIPAL
# ------------------------------

def main():
    print("=" * 70)
    print("INICIANDO PROCESSAMENTO DE INSUMOS")
    print("=" * 70)
    
    log_insumos = carregar_log(LOG_INSUMOS_FILE)
    
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
            
            # Enviar insumos
            log_insumos, sucessos, falhas = enviar_insumos_lote(df, log_insumos)
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