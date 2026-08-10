# test_connection.py
import os
from dotenv import load_dotenv
load_dotenv()

from graph_api import GraphClient, GraphConfig, test_sharepoint_connection, load_sharepoint_data
import pandas as pd

print("=== TESTE DE CONEXÃO SHAREPOINT ===\n")

# Verifica configuração
print(f"📋 Configurações:")
print(f"  Tenant: {GraphConfig.TENANT_ID[:8]}...")
print(f"  Client: {GraphConfig.CLIENT_ID[:8]}...")
print(f"  Site: {GraphConfig.SHAREPOINT_SITE_URL[:50]}...")
print(f"  Arquivo: {GraphConfig.EXCEL_FILENAME}")

# Testa conexão
success, message = test_sharepoint_connection()
print(f"\n🔗 Status: {'✅' if success else '❌'} {message}")

if success:
    print("\n📊 Carregando dados do Excel...")
    client = GraphClient()
    data = client.load_dashboard_data()
    
    df_controle = data.get("controle", pd.DataFrame())
    print(f"  Linhas carregadas: {len(df_controle)}")
    
    # Mostra os últimos valores
    if not df_controle.empty:
        print("\n📈 Últimas linhas:")
        print(df_controle.tail(3))
else:
    print("\n❌ Conexão falhou!")