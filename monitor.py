# monitor.py
import time
import pandas as pd
from datetime import datetime
from graph_api import GraphClient, invalidate_sharepoint_cache

print("=== MONITORANDO MUDANÇAS NO SHAREPOINT ===\n")
print("🔄 A cada 10 segundos vou verificar se a planilha mudou")
print("📝 Edite a planilha e salve, depois veja as mudanças aqui\n")
print("Pressione Ctrl+C para parar\n")

client = GraphClient()
ultima_linha = None

while True:
    try:
        # Força a recarga do cache
        invalidate_sharepoint_cache()
        
        # Carrega os dados
        data = client.load_dashboard_data()
        df = data.get("controle", pd.DataFrame())
        
        if not df.empty:
            # Mostra informações atuais
            print(f"\n[{datetime.now().strftime('%H:%M:%S')}]")
            print(f"  Total de linhas: {len(df)}")
            
            # Mostra as últimas 3 linhas
            print("\n  Últimas linhas:")
            for idx, row in df.tail(3).iterrows():
                veiculo = row.get('veiculo', '')
                total = row.get('total', 0)
                print(f"    {veiculo}: R$ {total:,.2f}")
            
            # Verifica se tem a linha "teste"
            teste_row = df[df['veiculo'].str.contains('teste', case=False, na=False)]
            if not teste_row.empty:
                print(f"\n  ✅ LINHA 'TESTE' ENCONTRADA!")
                print(f"    {teste_row.iloc[0]['veiculo']}: R$ {teste_row.iloc[0].get('total', 0):,.2f}")
            else:
                print(f"\n  ⚠️ Linha 'teste' NÃO encontrada")
        
        # Espera 10 segundos
        time.sleep(10)
        
    except KeyboardInterrupt:
        print("\n\n👋 Monitoramento encerrado")
        break
    except Exception as e:
        print(f"❌ Erro: {str(e)}")
        time.sleep(5)