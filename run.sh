# ============================================
# run.sh - Script para executar o dashboard
# ============================================
#!/bin/bash

# Instalar dependências
pip install -r requirements.txt

# Executar o dashboard
streamlit run app.py --server.port=8501 --server.address=0.0.0.0