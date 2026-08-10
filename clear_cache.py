# clear_cache.py
import streamlit as st
import shutil
import os
from pathlib import Path

print("=== LIMPANDO CACHES ===\n")

# 1. Limpa cache do Streamlit
try:
    st.cache_data.clear()
    print("✅ Cache do Streamlit limpo")
except:
    print("⚠️ Erro ao limpar cache do Streamlit")

# 2. Limpa cache do GraphClient
try:
    from graph_api import invalidate_sharepoint_cache
    invalidate_sharepoint_cache()
    print("✅ Cache do SharePoint limpo")
except:
    print("⚠️ Erro ao limpar cache do SharePoint")

# 3. Remove pastas __pycache__
pycache_dirs = list(Path(".").glob("**/__pycache__"))
for dir_path in pycache_dirs:
    try:
        shutil.rmtree(dir_path)
        print(f"✅ Removido: {dir_path}")
    except:
        print(f"⚠️ Erro ao remover: {dir_path}")

# 4. Remove arquivos .pyc
pyc_files = list(Path(".").glob("**/*.pyc"))
for file_path in pyc_files:
    try:
        file_path.unlink()
        print(f"✅ Removido: {file_path}")
    except:
        print(f"⚠️ Erro ao remover: {file_path}")

# 5. Remove cache do Streamlit
cache_dir = Path(os.path.expanduser("~/.cache/streamlit"))
if cache_dir.exists():
    try:
        shutil.rmtree(cache_dir)
        print(f"✅ Removido: {cache_dir}")
    except:
        print(f"⚠️ Erro ao remover: {cache_dir}")

print("\n✅ Todos os caches foram limpos!")
print("🚀 Agora rode: streamlit run app.py")