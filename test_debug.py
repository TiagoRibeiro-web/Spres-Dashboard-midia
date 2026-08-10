# test_debug.py
from graph_api import GraphClient, invalidate_sharepoint_cache
import pandas as pd

invalidate_sharepoint_cache()
client = GraphClient()
data = client.load_dashboard_data()
df = data.get("controle", pd.DataFrame())

print("🔍 Procurando por 'teste':")
teste_rows = df[df['veiculo'].str.contains('teste', case=False, na=False)]
if not teste_rows.empty:
    print("✅ Encontrado!")
    print(teste_rows[['veiculo', 'jan/27']])
else:
    print("❌ Não encontrado")
    print(f"Total de linhas: {len(df)}")
    print("Últimas 3 linhas:")
    print(df[['veiculo']].tail(3) if 'veiculo' in df.columns else df.tail(3))