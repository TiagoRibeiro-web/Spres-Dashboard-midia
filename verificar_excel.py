# verificar_excel.py
import pandas as pd
from graph_api import GraphClient, invalidate_sharepoint_cache

print("=== VERIFICANDO DADOS DO EXCEL ===\n")

# Força limpeza do cache
invalidate_sharepoint_cache()

# Carrega os dados
client = GraphClient()
data = client.load_dashboard_data()
df = data.get("controle", pd.DataFrame())

print(f"📊 Total de linhas: {len(df)}")
print(f"\n📋 Colunas: {list(df.columns)}\n")

# Procura por "teste"
teste_rows = df[df['veiculo'].str.contains('teste', case=False, na=False)]
if not teste_rows.empty:
    print("✅ LINHA 'TESTE' ENCONTRADA!")
    print(teste_rows)
else:
    print("❌ Linha 'teste' NÃO encontrada")
    print("\n🔍 Últimas 5 linhas do DataFrame:")
    print(df.tail(5))