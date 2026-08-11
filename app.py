# ============================================
# app.py - Dashboard Gestão Mídia Pro (SPRES EDITION - TECH LAYOUT)
# Layout tecnológico com glassmorphism, neon glows e animações sutis
# Mantém a paleta azul/amarelo e fonte Inter originais
# Dados fidedignos ao Excel: base_spres_projeto_refatorada.xlsx
# ============================================
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io
import base64
from datetime import datetime
import time
import logging
import os

# ============================================
# IMPORTAÇÃO DA INTEGRAÇÃO SHAREPOINT
# ============================================
try:
    from graph_api import (
        GraphClient,
        load_sharepoint_data,
        get_sharepoint_big_numbers,
        get_sharepoint_totais_mensais,
        test_sharepoint_connection,
        GraphConfig,
        invalidate_sharepoint_cache
    )
    SHAREPOINT_AVAILABLE = True
    logger = logging.getLogger(__name__)
except ImportError as e:
    SHAREPOINT_AVAILABLE = False
    print(f"⚠️ Módulo graph_api não encontrado: {str(e)}. Usando apenas dados estáticos.")

# ============================================
# IMPORTAÇÃO DA INTEGRAÇÃO EXCEL
# ============================================
try:
    from excel_integration import ExcelIntegration, mostrar_status_integracao, criar_interface_upload
    EXCEL_INTEGRATION_AVAILABLE = True
except ImportError as e:
    EXCEL_INTEGRATION_AVAILABLE = False
    print(f"⚠️ Módulo excel_integration não encontrado: {str(e)}")

# ============================================
# CARREGAR VARIÁVEIS DE AMBIENTE / SECRETS
# ============================================
# Carrega .env local (apenas para desenvolvimento)
try:
    from dotenv import load_dotenv
    load_dotenv()
except:
    pass

def get_secret(key, required=True):
    """Obtém uma variável do secrets.toml ou .env."""
    try:
        if hasattr(st, 'secrets') and key in st.secrets:
            return st.secrets[key]
    except:
        pass
    
    value = os.getenv(key)
    if value is not None:
        return value
    
    if required:
        st.error(f"❌ Configuração '{key}' não encontrada no secrets.toml ou .env")
        st.stop()
    return None

# ============================================
# AUTENTICAÇÃO (APENAS SECRETS)
# ============================================
USUARIOS = get_secret("USERS", required=True)

# Se USUARIOS for None, usa fallback
if USUARIOS is None:
    USUARIOS = {
        "admin": "spres2026",
        "gestao": "midia2026"
    }
    print("⚠️ Usando credenciais padrão (modo desenvolvimento)")

# Configurações do Azure
AZURE_TENANT_ID = get_secret("AZURE_TENANT_ID", required=True)
AZURE_CLIENT_ID = get_secret("AZURE_CLIENT_ID", required=True)
AZURE_CLIENT_SECRET = get_secret("AZURE_CLIENT_SECRET", required=True)
SHAREPOINT_SITE_URL = get_secret("SHAREPOINT_SITE_URL", required=True)
EXCEL_FILENAME = get_secret("EXCEL_FILENAME", required=False) or "base_spres_projeto_refatorada.xlsx"
SHAREPOINT_FILE_URL = get_secret("SHAREPOINT_FILE_URL", required=False)
CACHE_MINUTES = int(get_secret("CACHE_MINUTES", required=False) or "30")

# Verifica se está em modo Cloud
IS_CLOUD = hasattr(st, 'secrets') and 'AZURE_TENANT_ID' in st.secrets
if IS_CLOUD:
    print("☁️ Rodando no Streamlit Cloud")
else:
    print("💻 Rodando localmente")
    
    
st.set_page_config(
    page_title="Gestão Mídia Pro | Spres",
    page_icon="🍊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ============================================
# 0. PALETA DE CORES SPRES (fiel ao logo)
# ============================================
SPRES_BLUE = "#004B8D"          # azul corporativo do logo
SPRES_BLUE_DARK = "#00325F"     # azul profundo (header / títulos)
SPRES_BLUE_LIGHT = "#2E7DD1"    # azul claro (auxiliar de gráficos)
SPRES_BLUE_PALE = "#E6F0FA"     # fundo suave
SPRES_YELLOW = "#FFD600"        # amarelo vivo do logo
SPRES_YELLOW_DARK = "#DAB833"   # amarelo escuro (contraste)
SPRES_YELLOW_LIGHT = "#E9DE81ED"  # amarelo leve
SPRES_ORANGE = "#FF8A1E"        # laranja (destaque saldo/alerts)
SPRES_WHITE = "#FFFFFF"
SPRES_CREAM = "#FFFDF5"
SPRES_TEXT = "#0A1F35"
SPRES_TEXT_MUTED = "#5B6E80"
SPRES_BORDER = "#DDE5EE"

# Cores tecnológicas derivadas
TECH_CYAN = "#00E5FF"
TECH_GLOW_BLUE = "rgba(46, 125, 209, 0.15)"
TECH_GLOW_YELLOW = "rgba(255, 214, 0, 0.12)"
TECH_GLASS_BG = "rgba(255, 255, 255, 0.72)"
TECH_GLASS_BORDER = "rgba(0, 75, 141, 0.18)"
TECH_DARK_CARD = "#0C1B2E"

CHART_PALETTE = [SPRES_BLUE, SPRES_YELLOW, SPRES_ORANGE, SPRES_BLUE_LIGHT,
                 SPRES_YELLOW_DARK, "#1B63A8", "#FFCE00", "#0F3C70"]

# ============================================
# 1. CARREGAMENTO DE DADOS FIDEDIGNOS AO EXCEL
# ============================================
def carregar_dados_estaticos():
    """Carrega os dados diretamente do Excel (versão estática fidedigna)"""

    meses = ['fev/26', 'mar/26', 'abr/26', 'mai/26', 'jun/26',
             'jul/26', 'ago/26', 'set/26', 'out/26', 'nov/26', 'dez/26', 'jan/27']

    # ============================================================
    # ABA "GERAL" - Total por mês (linha 16 do Excel)
    # ============================================================
    dados_geral = {
        'fev/26': 7000,
        'mar/26': 56624,
        'abr/26': 42572,
        'mai/26': 43004.24,
        'jun/26': 62544.4,
        'jul/26': 22300,
        'ago/26': 18350,
        'set/26': 33565,
        'out/26': 43935,
        'nov/26': 58114.64,
        'dez/26': 26555.68,
        'jan/27': 4500,
        'total': 419064.96
    }

    # ============================================================
    # ABA "GERAL CONTROLE (2) NOVA" - Todos os 41 itens
    # ============================================================
    dados_controle = pd.DataFrame([
        # Itens principais (linhas 2-20 do Excel)
        {'veiculo': 'Rádio', 'canal': 'Difusora e Nova Brasil', 'obs': 'nenhuma observação',
         'fev/26': 0, 'mar/26': 5674, 'abr/26': 5622, 'mai/26': 0, 'jun/26': 0,
         'jul/26': 0, 'ago/26': 0, 'set/26': 0, 'out/26': 0, 'nov/26': 0, 'dez/26': 0, 'jan/27': 0},

        {'veiculo': 'Rádio', 'canal': '106 e conquista', 'obs': 'nenhuma observação',
         'fev/26': 0, 'mar/26': 0, 'abr/26': 0, 'mai/26': 8954.24, 'jun/26': 9794.4,
         'jul/26': 0, 'ago/26': 0, 'set/26': 0, 'out/26': 0, 'nov/26': 0, 'dez/26': 0, 'jan/27': 0},

        {'veiculo': 'Rádio', 'canal': 'Kiss e Mix', 'obs': 'nenhuma observação',
         'fev/26': 0, 'mar/26': 0, 'abr/26': 0, 'mai/26': 0, 'jun/26': 0,
         'jul/26': 6600, 'ago/26': 5550, 'set/26': 0, 'out/26': 0, 'nov/26': 0, 'dez/26': 0, 'jan/27': 0},

        {'veiculo': 'Rádio', 'canal': 'Melody e Jovem Pan', 'obs': 'nenhuma observação',
         'fev/26': 0, 'mar/26': 0, 'abr/26': 0, 'mai/26': 0, 'jun/26': 0,
         'jul/26': 0, 'ago/26': 0, 'set/26': 6765, 'out/26': 6985, 'nov/26': 0, 'dez/26': 0, 'jan/27': 0},

        {'veiculo': 'Rádio', 'canal': 'Mega e Clube', 'obs': 'nenhuma observação',
         'fev/26': 0, 'mar/26': 0, 'abr/26': 0, 'mai/26': 0, 'jun/26': 0,
         'jul/26': 0, 'ago/26': 0, 'set/26': 0, 'out/26': 0, 'nov/26': 10064.64, 'dez/26': 10855.68, 'jan/27': 0},

        {'veiculo': 'Painel Rodovia', 'canal': 'EPTV', 'obs': 'nenhuma observação',
         'fev/26': 4500, 'mar/26': 4500, 'abr/26': 4500, 'mai/26': 4500, 'jun/26': 4500,
         'jul/26': 4500, 'ago/26': 4500, 'set/26': 4500, 'out/26': 4500, 'nov/26': 4500, 'dez/26': 4500, 'jan/27': 4500},

        {'veiculo': 'Outdoor - Pacote Nóbile', 'canal': '2 Bi semanas lançamento', 'obs': 'nenhuma observação',
         'fev/26': 0, 'mar/26': 21250, 'abr/26': 21250, 'mai/26': 21250, 'jun/26': 0,
         'jul/26': 0, 'ago/26': 0, 'set/26': 0, 'out/26': 21250, 'nov/26': 21250, 'dez/26': 0, 'jan/27': 0},

        {'veiculo': 'Outdoor - Pacote Nóbile', 'canal': '2 Bi semana Lanc. de Niver', 'obs': 'nenhuma observação',
         'fev/26': 0, 'mar/26': 0, 'abr/26': 0, 'mai/26': 0, 'jun/26': 21250,
         'jul/26': 0, 'ago/26': 0, 'set/26': 0, 'out/26': 0, 'nov/26': 0, 'dez/26': 0, 'jan/27': 0},

        {'veiculo': 'Painel de Led', 'canal': 'Produto e Lançamento sabor', 'obs': 'usado em março abril (março)',
         'fev/26': 0, 'mar/26': 14000, 'abr/26': 0, 'mai/26': 0, 'jun/26': 0,
         'jul/26': 0, 'ago/26': 0, 'set/26': 0, 'out/26': 0, 'nov/26': 0, 'dez/26': 0, 'jan/27': 0},

        {'veiculo': 'Painel de Led', 'canal': 'Niver e Copa', 'obs': 'usado em junho e julho (junho)',
         'fev/26': 0, 'mar/26': 0, 'abr/26': 0, 'mai/26': 0, 'jun/26': 14000,
         'jul/26': 0, 'ago/26': 0, 'set/26': 0, 'out/26': 0, 'nov/26': 0, 'dez/26': 0, 'jan/27': 0},

        {'veiculo': 'Painel de Led', 'canal': 'Institucional', 'obs': 'usado em setembro e outubro (Setembro)',
         'fev/26': 0, 'mar/26': 0, 'abr/26': 0, 'mai/26': 0, 'jun/26': 0,
         'jul/26': 0, 'ago/26': 0, 'set/26': 14000, 'out/26': 0, 'nov/26': 0, 'dez/26': 0, 'jan/27': 0},

        {'veiculo': 'Painel de Led', 'canal': 'Institucional', 'obs': 'usado em novembro e dezembro (novembro)',
         'fev/26': 0, 'mar/26': 0, 'abr/26': 0, 'mai/26': 0, 'jun/26': 0,
         'jul/26': 0, 'ago/26': 0, 'set/26': 0, 'out/26': 0, 'nov/26': 14000, 'dez/26': 0, 'jan/27': 0},

        {'veiculo': 'Campanha Digital', 'canal': 'Mídia Digital', 'obs': 'nenhuma observação',
         'fev/26': 0, 'mar/26': 7200, 'abr/26': 7200, 'mai/26': 4300, 'jun/26': 9000,
         'jul/26': 7200, 'ago/26': 4300, 'set/26': 4300, 'out/26': 7200, 'nov/26': 4300, 'dez/26': 7200, 'jan/27': 0},

        {'veiculo': 'Influencers + Chef Spres', 'canal': 'Mídia Digital', 'obs': 'nenhuma observação',
         'fev/26': 0, 'mar/26': 4000, 'abr/26': 4000, 'mai/26': 4000, 'jun/26': 4000,
         'jul/26': 4000, 'ago/26': 4000, 'set/26': 4000, 'out/26': 4000, 'nov/26': 4000, 'dez/26': 4000, 'jan/27': 0},

        {'veiculo': 'Chef Spress', 'canal': 'Mídia Digital', 'obs': 'nenhuma observação',
         'fev/26': 2500, 'mar/26': 400, 'abr/26': 0, 'mai/26': 0, 'jun/26': 0,
         'jul/26': 0, 'ago/26': 0, 'set/26': 0, 'out/26': 0, 'nov/26': 0, 'dez/26': 0, 'jan/27': 0},

        {'veiculo': 'Troféu Copa Spres', 'canal': 'Mídia', 'obs': 'nenhuma observação',
         'fev/26': 0, 'mar/26': 480, 'abr/26': 0, 'mai/26': 0, 'jun/26': 0,
         'jul/26': 0, 'ago/26': 0, 'set/26': 0, 'out/26': 0, 'nov/26': 0, 'dez/26': 0, 'jan/27': 0},

        {'veiculo': 'Carta Influenciador Brunch Run', 'canal': 'Mídia', 'obs': 'nenhuma observação',
         'fev/26': 0, 'mar/26': 34.5, 'abr/26': 0, 'mai/26': 0, 'jun/26': 0,
         'jul/26': 0, 'ago/26': 0, 'set/26': 0, 'out/26': 0, 'nov/26': 0, 'dez/26': 0, 'jan/27': 0},

        {'veiculo': 'Voucher Brunch Run', 'canal': 'Mídia', 'obs': 'nenhuma observação',
         'fev/26': 0, 'mar/26': 506, 'abr/26': 0, 'mai/26': 0, 'jun/26': 0,
         'jul/26': 0, 'ago/26': 0, 'set/26': 0, 'out/26': 0, 'nov/26': 0, 'dez/26': 0, 'jan/27': 0},

        {'veiculo': '1º Pacela Laricas', 'canal': 'Mídia', 'obs': 'nenhuma observação',
         'fev/26': 0, 'mar/26': 0, 'abr/26': 0, 'mai/26': 0, 'jun/26': 2250,
         'jul/26': 0, 'ago/26': 0, 'set/26': 0, 'out/26': 0, 'nov/26': 0, 'dez/26': 0, 'jan/27': 0},

        {'veiculo': 'Artesano Collors Run (saldo)', 'canal': '', 'obs': 'debitado em março',
         'fev/26': 0, 'mar/26': 1000, 'abr/26': 1000, 'mai/26': 0, 'jun/26': 0,
         'jul/26': 0, 'ago/26': 0, 'set/26': 0, 'out/26': 0, 'nov/26': 0, 'dez/26': 0, 'jan/27': 0},

        {'veiculo': 'Copa Ribeirão Beach Tenis (saldo)', 'canal': '', 'obs': 'Sobrou de março (debitado em abril)',
         'fev/26': 0, 'mar/26': 0, 'abr/26': 1500, 'mai/26': 1500, 'jun/26': 0,
         'jul/26': 0, 'ago/26': 0, 'set/26': 0, 'out/26': 0, 'nov/26': 0, 'dez/26': 0, 'jan/27': 0},

        # Itens extras do Excel (linhas 23-42)
        {'veiculo': 'Kit Influenciador - Sacolas', 'canal': '', 'obs': '',
         'fev/26': 0, 'mar/26': 0, 'abr/26': 188, 'mai/26': 0, 'jun/26': 0,
         'jul/26': 0, 'ago/26': 0, 'set/26': 0, 'out/26': 0, 'nov/26': 0, 'dez/26': 0, 'jan/27': 0},

        {'veiculo': 'Kit Influenciador - Tag', 'canal': '', 'obs': '',
         'fev/26': 0, 'mar/26': 0, 'abr/26': 161, 'mai/26': 0, 'jun/26': 0,
         'jul/26': 0, 'ago/26': 0, 'set/26': 0, 'out/26': 0, 'nov/26': 0, 'dez/26': 0, 'jan/27': 0},

        {'veiculo': 'Kit Influenciador - Frete', 'canal': '', 'obs': '',
         'fev/26': 0, 'mar/26': 0, 'abr/26': 195, 'mai/26': 0, 'jun/26': 0,
         'jul/26': 0, 'ago/26': 0, 'set/26': 0, 'out/26': 0, 'nov/26': 0, 'dez/26': 0, 'jan/27': 0},

        {'veiculo': 'Kit Influenciador - Carta', 'canal': '', 'obs': '',
         'fev/26': 0, 'mar/26': 0, 'abr/26': 50, 'mai/26': 0, 'jun/26': 0,
         'jul/26': 0, 'ago/26': 0, 'set/26': 0, 'out/26': 0, 'nov/26': 0, 'dez/26': 0, 'jan/27': 0},

        {'veiculo': 'Caixa Lançamento', 'canal': '', 'obs': '',
         'fev/26': 0, 'mar/26': 0, 'abr/26': 460, 'mai/26': 0, 'jun/26': 0,
         'jul/26': 0, 'ago/26': 0, 'set/26': 0, 'out/26': 0, 'nov/26': 0, 'dez/26': 0, 'jan/27': 0},

        {'veiculo': 'Voucher 15% OFF Copa Beach Tenis', 'canal': '', 'obs': '',
         'fev/26': 0, 'mar/26': 0, 'abr/26': 302, 'mai/26': 0, 'jun/26': 0,
         'jul/26': 0, 'ago/26': 0, 'set/26': 0, 'out/26': 0, 'nov/26': 0, 'dez/26': 0, 'jan/27': 0},

        {'veiculo': 'Captação Video', 'canal': '', 'obs': '',
         'fev/26': 0, 'mar/26': 0, 'abr/26': 580, 'mai/26': 0, 'jun/26': 0,
         'jul/26': 0, 'ago/26': 0, 'set/26': 0, 'out/26': 0, 'nov/26': 0, 'dez/26': 0, 'jan/27': 0},

        {'veiculo': 'Promotora Lançamento', 'canal': '', 'obs': '',
         'fev/26': 0, 'mar/26': 0, 'abr/26': 223.35, 'mai/26': 0, 'jun/26': 0,
         'jul/26': 0, 'ago/26': 0, 'set/26': 0, 'out/26': 0, 'nov/26': 0, 'dez/26': 0, 'jan/27': 0},

        {'veiculo': 'Copos Sede de Torcer', 'canal': '', 'obs': '',
         'fev/26': 0, 'mar/26': 0, 'abr/26': 0, 'mai/26': 0, 'jun/26': 0,
         'jul/26': 0, 'ago/26': 0, 'set/26': 0, 'out/26': 0, 'nov/26': 0, 'dez/26': 0, 'jan/27': 0},

        {'veiculo': 'Camiseta Laricas', 'canal': '', 'obs': 'nenhuma observação',
         'fev/26': 0, 'mar/26': 0, 'abr/26': 0, 'mai/26': 799, 'jun/26': 0,
         'jul/26': 0, 'ago/26': 0, 'set/26': 0, 'out/26': 0, 'nov/26': 0, 'dez/26': 0, 'jan/27': 0},

        {'veiculo': 'Camiseta Laricas 2', 'canal': '', 'obs': 'nenhuma observação',
         'fev/26': 0, 'mar/26': 0, 'abr/26': 0, 'mai/26': 0, 'jun/26': 572,
         'jul/26': 0, 'ago/26': 0, 'set/26': 0, 'out/26': 0, 'nov/26': 0, 'dez/26': 0, 'jan/27': 0},

        {'veiculo': 'Avental Laricas', 'canal': '', 'obs': 'nenhuma observação',
         'fev/26': 0, 'mar/26': 0, 'abr/26': 0, 'mai/26': 192, 'jun/26': 0,
         'jul/26': 0, 'ago/26': 0, 'set/26': 0, 'out/26': 0, 'nov/26': 0, 'dez/26': 0, 'jan/27': 0},

        {'veiculo': 'Microfone Laricas', 'canal': '', 'obs': 'nenhuma observação',
         'fev/26': 0, 'mar/26': 0, 'abr/26': 0, 'mai/26': 150.89, 'jun/26': 0,
         'jul/26': 0, 'ago/26': 0, 'set/26': 0, 'out/26': 0, 'nov/26': 0, 'dez/26': 0, 'jan/27': 0},

        {'veiculo': 'Vibra Run', 'canal': '', 'obs': '',
         'fev/26': 0, 'mar/26': 0, 'abr/26': 0, 'mai/26': 0, 'jun/26': 0,
         'jul/26': 0, 'ago/26': 0, 'set/26': 0, 'out/26': 0, 'nov/26': 0, 'dez/26': 0, 'jan/27': 0},

        {'veiculo': '2º Pacela Laricas', 'canal': '', 'obs': '',
         'fev/26': 0, 'mar/26': 0, 'abr/26': 0, 'mai/26': 0, 'jun/26': 0,
         'jul/26': 2250, 'ago/26': 0, 'set/26': 0, 'out/26': 0, 'nov/26': 0, 'dez/26': 0, 'jan/27': 0},

        {'veiculo': 'Sacolas Apas NH', 'canal': '', 'obs': '',
         'fev/26': 0, 'mar/26': 0, 'abr/26': 0, 'mai/26': 0, 'jun/26': 0,
         'jul/26': 448.5, 'ago/26': 0, 'set/26': 0, 'out/26': 0, 'nov/26': 0, 'dez/26': 0, 'jan/27': 0},

        {'veiculo': 'Avental Chef Spres', 'canal': '', 'obs': '',
         'fev/26': 0, 'mar/26': 0, 'abr/26': 0, 'mai/26': 0, 'jun/26': 0,
         'jul/26': 0, 'ago/26': 310.5, 'set/26': 0, 'out/26': 0, 'nov/26': 0, 'dez/26': 0, 'jan/27': 0},

        {'veiculo': 'Flyer Corrida Tribuna', 'canal': '', 'obs': '',
         'fev/26': 0, 'mar/26': 0, 'abr/26': 0, 'mai/26': 0, 'jun/26': 0,
         'jul/26': 674, 'ago/26': 0, 'set/26': 0, 'out/26': 0, 'nov/26': 0, 'dez/26': 0, 'jan/27': 0},

        {'veiculo': 'Patrocinio Corrida Brunch Run', 'canal': '', 'obs': '',
         'fev/26': 0, 'mar/26': 0, 'abr/26': 0, 'mai/26': 0, 'jun/26': 0,
         'jul/26': 0, 'ago/26': 0, 'set/26': 0, 'out/26': 0, 'nov/26': 0, 'dez/26': 0, 'jan/27': 0},

        {'veiculo': 'Patrocinio Aulão Doha Sports', 'canal': '', 'obs': '',
         'fev/26': 0, 'mar/26': 0, 'abr/26': 0, 'mai/26': 0, 'jun/26': 0,
         'jul/26': 350, 'ago/26': 0, 'set/26': 0, 'out/26': 0, 'nov/26': 0, 'dez/26': 0, 'jan/27': 0},
    ])

    # ============================================================
    # BIG NUMBERS - Fidedignos ao Excel (linha 47-49 da aba Geral Controle (2) nova)
    # ============================================================
    # Linha 47: Total Controle Aprovado = 419.835,46
    # Linha 48: Total Controle Utilizado = 244.971,38 (somando fev a ago)
    # Linha 49: Total Saldo Positivo = 7.423,26 (somando fev a ago)
    big_numbers = {
        'total_controle_aprovado': 419835.46,
        'total_controle_utilizado': 244971.38,
        'total_saldo_positivo': 7423.26
    }

    # Totais mensais do Controle Aprovado (linha 47)
    totais_aprovado_mensal = {
        'fev/26': 7000, 'mar/26': 56624, 'abr/26': 42572, 'mai/26': 43004.24,
        'jun/26': 62544.4, 'jul/26': 22300, 'ago/26': 18350, 'set/26': 33565,
        'out/26': 39935, 'nov/26': 58114.64, 'dez/26': 26555.68, 'jan/27': 4500
    }

    # Totais mensais do Controle Utilizado (linha 48)
    totais_utilizado_mensal = {
        'fev/26': 7000, 'mar/26': 55044.5, 'abr/26': 43231.35, 'mai/26': 41646.13,
        'jun/26': 61366.4, 'jul/26': 22022.5, 'ago/26': 14660.5, 'set/26': 0,
        'out/26': 0, 'nov/26': 0, 'dez/26': 0, 'jan/27': 0
    }

    # Totais mensais do Saldo Positivo (linha 49)
    totais_saldo_mensal = {
        'fev/26': 0, 'mar/26': 1579.5, 'abr/26': -659.35, 'mai/26': 1358.11,
        'jun/26': 1178, 'jul/26': 277.5, 'ago/26': 3689.5, 'set/26': 0,
        'out/26': 0, 'nov/26': 0, 'dez/26': 0, 'jan/27': 0
    }

    # ============================================================
    # ABA "ESTUDO" - Dados para comparação (mantido para compatibilidade)
    # ============================================================
    dados_estudo = {
        'aba_geral': {},
        'aba_geral_controle': {},
        'aba_geral_extra': {}
    }

    # ============================================================
    # ABA "GERAL EXTRA" - Itens extras (mantido para compatibilidade)
    # ============================================================
    dados_extra = pd.DataFrame()

    return (dados_geral, dados_controle, big_numbers, totais_aprovado_mensal,
            totais_utilizado_mensal, totais_saldo_mensal, dados_estudo, dados_extra)

# ============================================
# 1.1. CARREGAMENTO DE DADOS DO SHAREPOINT (NOVO)
# ============================================
def carregar_dados_sharepoint():
    """
    Carrega dados do SharePoint em tempo real.
    Se falhar, usa fallback para dados estáticos.
    """
    if not SHAREPOINT_AVAILABLE:
        logger.warning("Módulo SharePoint não disponível. Usando dados estáticos.")
        return carregar_dados_estaticos()
    
    if not GraphConfig.is_configured():
        logger.warning("Configuração do SharePoint incompleta. Usando dados estáticos.")
        return carregar_dados_estaticos()
    
    try:
        with st.spinner("🔄 Carregando dados do SharePoint em tempo real..."):
            invalidate_sharepoint_cache()
            st.cache_data.clear()
            data = load_sharepoint_data()
            
            if "error" in data:
                st.warning(f"⚠️ Erro ao carregar do SharePoint: {data['error']}")
                st.info("Usando dados estáticos como fallback.")
                return carregar_dados_estaticos()
            
            return processar_dados_sharepoint(data)
            
    except Exception as e:
        logger.error(f"Erro na integração SharePoint: {str(e)}")
        st.warning(f"⚠️ Erro ao conectar com SharePoint: {str(e)}")
        st.info("Usando dados estáticos como fallback.")
        return carregar_dados_estaticos()

def processar_dados_sharepoint(data: dict):
    """
    Processa os dados do SharePoint para o formato do dashboard.
    CORRIGIDO: Conversão de valores com formato brasileiro (R$ 1.234,56)
    """
    meses = ['fev/26', 'mar/26', 'abr/26', 'mai/26', 'jun/26',
             'jul/26', 'ago/26', 'set/26', 'out/26', 'nov/26', 'dez/26', 'jan/27']
    
    # Extrai DataFrame de controle
    df_controle = data.get("controle", pd.DataFrame())
    
    if df_controle.empty:
        logger.warning("Dados de controle vazios do SharePoint. Usando estático.")
        return carregar_dados_estaticos()
    
    # ============================================================
    # FUNÇÃO DE CONVERSÃO CORRIGIDA
    # ============================================================
    def converter_valor(x):
        """
        Converte valor monetário para float.
        Suporta: "R$ 7.000,00" -> 7000.0, "R$ 56.624,00" -> 56624.0
        """
        if pd.isna(x) or x == '' or x == '-' or x is None:
            return 0.0
        if isinstance(x, (int, float)):
            return float(x)
        if isinstance(x, str):
            # Remove "R$", espaços
            valor_limpo = x.replace('R$', '').replace(' ', '').strip()
            # Remove pontos de milhar (.) e substitui vírgula (,) por ponto (.)
            # Exemplo: "7.000,00" -> "7000.00"
            valor_limpo = valor_limpo.replace('.', '').replace(',', '.')
            # Remove qualquer caractere que não seja número ou ponto
            import re
            valor_limpo = re.sub(r'[^\d.]', '', valor_limpo)
            try:
                return float(valor_limpo) if valor_limpo else 0.0
            except:
                return 0.0
        return 0.0
    
    # ============================================================
    # IDENTIFICA AS COLUNAS
    # ============================================================
    # Verifica se as colunas já estão padronizadas
    if 'veiculo' in df_controle.columns:
        # Já está padronizado, usa como está
        pass
    else:
        # Tenta identificar as colunas pela posição
        colunas_df = list(df_controle.columns)
        
        # Mapeamento baseado na posição
        mapeamento = {}
        
        # Primeira coluna = veiculo
        if len(colunas_df) > 0:
            mapeamento[colunas_df[0]] = 'veiculo'
        # Segunda coluna = canal
        if len(colunas_df) > 1:
            mapeamento[colunas_df[1]] = 'canal'
        # Terceira coluna = obs
        if len(colunas_df) > 2:
            mapeamento[colunas_df[2]] = 'obs'
        
        # Identifica colunas de meses
        for col in colunas_df:
            col_str = str(col).lower()
            for mes in meses:
                if mes.lower() in col_str or col_str in mes.lower():
                    mapeamento[col] = mes
                    break
        
        # Última coluna = total (se não for um mês)
        if len(colunas_df) > 0:
            ultima_coluna = colunas_df[-1]
            if ultima_coluna not in mapeamento:
                mapeamento[ultima_coluna] = 'total'
        
        # Aplica o mapeamento
        df_controle = df_controle.rename(columns=mapeamento)
    
    # ============================================================
    # GARANTE QUE TODOS OS MESES EXISTEM
    # ============================================================
    for mes in meses:
        if mes not in df_controle.columns:
            df_controle[mes] = 0.0
    
    # ============================================================
    # CONVERTE TODOS OS MESES PARA NUMÉRICO
    # ============================================================
    for mes in meses:
        if mes in df_controle.columns:
            df_controle[mes] = df_controle[mes].apply(converter_valor)
    
    # ============================================================
    # REMOVE LINHAS DE TOTAL (para não contaminar os gráficos)
    # ============================================================
    if 'veiculo' in df_controle.columns:
        # Remove linhas que contêm "Total", "Saldo", etc.
        mask_total = df_controle['veiculo'].astype(str).str.contains(
            'Total|TOTAL|Saldo|SALDO|Aprovado|Utilizado|Positivo', 
            case=False, 
            na=False
        )
        df_controle = df_controle[~mask_total]
    
    # ============================================================
    # EXTRAI DADOS GERAIS (soma por mês)
    # ============================================================
    dados_geral = {}
    total_geral = 0
    
    for mes in meses:
        if mes in df_controle.columns:
            valor = df_controle[mes].sum()
            dados_geral[mes] = float(valor) if not pd.isna(valor) else 0
            total_geral += dados_geral[mes]
        else:
            dados_geral[mes] = 0
    
    dados_geral['total'] = total_geral
    
    # ============================================================
    # CALCULA TOTAL POR LINHA
    # ============================================================
    if 'total' not in df_controle.columns:
        df_controle['total'] = df_controle[meses].sum(axis=1)
    
    # ============================================================
    # BIG NUMBERS
    # ============================================================
    try:
        big_numbers = get_sharepoint_big_numbers()
    except:
        big_numbers = {
            'total_controle_aprovado': 0,
            'total_controle_utilizado': 0,
            'total_saldo_positivo': 0
        }
    
    # ============================================================
    # TOTAIS MENSAIS
    # ============================================================
    try:
        totais_mensais = get_sharepoint_totais_mensais()
        totais_aprovado_mensal = totais_mensais.get('aprovado', {})
        totais_utilizado_mensal = totais_mensais.get('utilizado', {})
        totais_saldo_mensal = totais_mensais.get('saldo', {})
    except:
        totais_aprovado_mensal = {}
        totais_utilizado_mensal = {}
        totais_saldo_mensal = {}
    
    if not totais_aprovado_mensal:
        totais_aprovado_mensal = {mes: 0 for mes in meses}
        totais_utilizado_mensal = {mes: 0 for mes in meses}
        totais_saldo_mensal = {mes: 0 for mes in meses}
    
    # ============================================================
    # DADOS DE ESTUDO E EXTRA (vazios para compatibilidade)
    # ============================================================
    dados_estudo = {
        'aba_geral': {},
        'aba_geral_controle': {},
        'aba_geral_extra': {}
    }
    dados_extra = pd.DataFrame()
    
    return (dados_geral, df_controle, big_numbers, totais_aprovado_mensal,
            totais_utilizado_mensal, totais_saldo_mensal, dados_estudo, dados_extra)

# ============================================
# 2. FUNÇÕES DE PROCESSAMENTO
# ============================================
def processar_dados(dados_controle):
    """Processa os dados para visualizações - CORRIGIDO"""
    
    meses = ['fev/26', 'mar/26', 'abr/26', 'mai/26', 'jun/26',
             'jul/26', 'ago/26', 'set/26', 'out/26', 'nov/26', 'dez/26', 'jan/27']
    
    # ===== FAZ UMA CÓPIA PARA NÃO MODIFICAR O ORIGINAL =====
    df = dados_controle.copy()
    # ===== REMOVE APENAS LINHAS DE TOTAL, NÃO LINHAS COM "teste" =====
    if 'veiculo' in dados_controle.columns:
        # Remove apenas linhas que contêm "Total" ou "Saldo"
        dados_controle = dados_controle[~dados_controle['veiculo'].str.contains('Total|TOTAL|Saldo|SALDO', case=False, na=False)]
        # NÃO remove "teste"
    # ===== REMOVE LINHAS DE TOTAL =====
    if 'veiculo' in df.columns:
        df = df[~df['veiculo'].str.contains('Total|TOTAL|Saldo|SALDO', case=False, na=False)]
        df = df[~df['veiculo'].str.contains('Total', case=False, na=False)]
    
    # ===== GARANTE QUE TODOS OS MESES EXISTEM =====
    for mes in meses:
        if mes not in df.columns:
            df[mes] = 0.0
    
    # ===== CONVERTE VALORES PARA NUMÉRICO (CORRIGIDO) =====
    for mes in meses:
        def converter(x):
            if pd.isna(x) or x == '' or x == '-' or x is None:
                return 0.0
            if isinstance(x, (int, float)):
                return float(x)
            if isinstance(x, str):
                # Remove "R$", espaços, pontos de milhar
                # Exemplo: "R$ 7.000,00" -> "7000.00"
                valor_limpo = x.replace('R$', '').replace(' ', '').replace('.', '').replace(',', '.')
                import re
                valor_limpo = re.sub(r'[^\d.]', '', valor_limpo)
                try:
                    return float(valor_limpo) if valor_limpo else 0.0
                except:
                    return 0.0
            return 0.0
        
        df[mes] = df[mes].apply(converter)
    
    # ===== CALCULA TOTAL POR LINHA =====
    df['total'] = df[meses].sum(axis=1)
    
    # ===== CALCULA TOTAL POR MÊS =====
    dados_mensais = df[meses].sum()
    
    # ===== DISTRIBUIÇÃO POR VEÍCULO =====
    distribuicao_veiculo = df.groupby('veiculo')['total'].sum().sort_values(ascending=True)
    
    return df, distribuicao_veiculo, dados_mensais

# ============================================
# 3. FUNÇÕES DE VISUALIZAÇÃO (COM BIG NUMBERS)
# ============================================
def formatar_moeda(valor):
    """Formata número para moeda brasileira."""
    if valor is None:
        return '-'
    return f"R$ {valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')


def criar_big_numbers(big_numbers, totais_aprovado_mensal, totais_utilizado_mensal, totais_saldo_mensal):
    """Cria os big numbers destacados com identidade Spres tecnológica"""

    total_aprovado = big_numbers['total_controle_aprovado']
    total_utilizado = big_numbers['total_controle_utilizado']
    saldo_positivo = big_numbers['total_saldo_positivo']

    # % de execução
    pct_utilizado = (total_utilizado / total_aprovado * 100) if total_aprovado > 0 else 0
    pct_saldo = (saldo_positivo / total_aprovado * 100) if total_aprovado > 0 else 0

    html = f"""
    <div class="big-number-container">
        <div class="big-number-item tech-glow-blue">
            <div class="big-number-header">
                <span class="big-number-icon-tech">✓</span>
            </div>
            <div class="big-number-label">Total Controle Aprovado</div>
            <div class="big-number-value" style="color:{SPRES_BLUE};">{formatar_moeda(total_aprovado)}</div>
            <div class="big-number-sub">Orçamento aprovado (linha 47)</div>
            <div class="big-number-bar">
                <div class="big-number-bar-fill" style="width:100%; background: {SPRES_BLUE};"></div>
            </div>
        </div>
        <div class="big-number-item tech-glow-yellow">
            <div class="big-number-header">
                <span class="big-number-icon-tech">▶</span>
                <span class="big-number-pct">{pct_utilizado:.1f}%</span>
            </div>
            <div class="big-number-label">Total Controle Utilizado</div>
            <div class="big-number-value" style="color:{SPRES_YELLOW_DARK};">{formatar_moeda(total_utilizado)}</div>
            <div class="big-number-sub">Gastos realizados (linha 48)</div>
            <div class="big-number-bar">
                <div class="big-number-bar-fill" style="width:{pct_utilizado:.1f}%; background: {SPRES_YELLOW_DARK};"></div>
            </div>
        </div>
        <div class="big-number-item tech-glow-orange">
            <div class="big-number-header">
                <span class="big-number-icon-tech">◆</span>
                <span class="big-number-pct">{pct_saldo:.1f}%</span>
            </div>
            <div class="big-number-label">Total Saldo Positivo</div>
            <div class="big-number-value" style="color:{SPRES_ORANGE};">{formatar_moeda(saldo_positivo)}</div>
            <div class="big-number-sub">Disponível para uso (linha 49)</div>
            <div class="big-number-bar">
                <div class="big-number-bar-fill" style="width:{pct_saldo:.1f}%; background: {SPRES_ORANGE};"></div>
            </div>
        </div>
    </div>
    """

    st.markdown(html, unsafe_allow_html=True)

def criar_cards(dados_geral, dados_controle):
    """Cria os cards de métricas principais"""

    total_geral = dados_geral['total']
    total_controle = dados_controle['total'].sum()
    saldo = total_geral - total_controle

    col1, col2, col3 = st.columns(3)

    with col1:
        st.metric(label="💰 Total Geral", value=formatar_moeda(total_geral))
    with col2:
        st.metric(label="📋 Controle Aprovado", value=formatar_moeda(total_controle))
    with col3:
        st.metric(label="� Saldo", value=formatar_moeda(saldo),
            delta=formatar_moeda(abs(saldo)) if saldo >= 0 else f"-{formatar_moeda(abs(saldo))}"
        )

def criar_grafico_mensal(dados_mensais):
    """Cria gráfico de barras do investimento mensal com identidade Spres tecnológica"""

    meses = ['Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez', 'Jan']
    
    # Converte para lista e garante que todos os valores são numéricos
    if hasattr(dados_mensais, 'values'):
        valores = list(dados_mensais.values)
    else:
        valores = list(dados_mensais)
    
    # Garante que temos 12 meses
    if len(valores) < 12:
        valores = valores + [0] * (12 - len(valores))
    
    # Remove valores NaN, None, infinitos e converte para float
    import math
    valores_limpos = []
    for v in valores[:12]:  # Pega apenas os primeiros 12
        try:
            if v is None or (isinstance(v, float) and (math.isnan(v) or math.isinf(v))):
                valores_limpos.append(0.0)
            else:
                valores_limpos.append(float(v))
        except (ValueError, TypeError):
            valores_limpos.append(0.0)
    
    # Cria textos formatados para as barras
    textos = []
    for v in valores_limpos:
        if v > 0:
            textos.append(f'R$ {v:,.0f}'.replace(',', '.'))
        else:
            textos.append('')
    
    # Cria o gráfico
    fig = go.Figure(data=[
        go.Bar(
            x=meses,
            y=valores_limpos,
            marker=dict(
                color=valores_limpos,
                colorscale=[
                    [0, SPRES_BLUE_LIGHT],
                    [0.5, SPRES_BLUE],
                    [1, SPRES_BLUE_DARK]
                ],
                line=dict(color=SPRES_YELLOW, width=1.2),
                cornerradius=6
            ),
            text=textos,
            textposition='outside',
            textfont=dict(size=10, color=SPRES_TEXT, family='Inter, sans-serif'),
            hovertemplate='<b>%{x}</b><br>Investimento: R$ %{y:,.2f}<extra></extra>',
            showlegend=False
        )
    ])

    fig.update_layout(
        title=None,
        xaxis_title=None,
        yaxis_title=None,
        height=360,
        margin=dict(l=20, r=20, t=30, b=20),
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(color=SPRES_TEXT, family='Inter, sans-serif'),
        showlegend=False,
        xaxis=dict(
            showgrid=False,
            linecolor=SPRES_BORDER,
            tickfont=dict(size=12, color=SPRES_TEXT_MUTED)
        ),
        yaxis=dict(
            gridcolor='rgba(0,75,141,0.08)',
            showgrid=True,
            zerolinecolor=SPRES_BORDER,
            tickprefix='R$ ',
            tickfont=dict(size=11, color=SPRES_TEXT_MUTED)
        )
    )

    return fig

def criar_grafico_distribuicao(distribuicao_veiculo, total_geral=None):
    """
    Cria treemap com EXATAMENTE 6-7 retângulos
    Cores 100% Spres: azul corporativo, amarelo, laranja
    Porcentagens calculadas sobre o total geral
    """
    
    # Remove valores nulos ou vazios
    distribuicao_veiculo = distribuicao_veiculo[distribuicao_veiculo.index.notna()]
    distribuicao_veiculo = distribuicao_veiculo[distribuicao_veiculo.index != '']
    distribuicao_veiculo = distribuicao_veiculo[distribuicao_veiculo > 0]
    
    if distribuicao_veiculo.empty:
        fig = go.Figure()
        fig.add_annotation(text="Sem dados para exibir", x=0.5, y=0.5, showarrow=False)
        fig.update_layout(height=420)
        return fig

    # Usa o total geral se fornecido
    if total_geral is None:
        total_geral = distribuicao_veiculo.sum()
    
    # Ordena por valor decrescente
    distribuicao_ord = distribuicao_veiculo.sort_values(ascending=False)
    
    # Pega os TOP 5 principais
    top5 = distribuicao_ord.head(5)
    
    # O resto vira "Outros"
    outros_valor = distribuicao_ord.iloc[5:].sum() if len(distribuicao_ord) > 5 else 0
    
    # Monta os labels e valores (5 principais + Outros)
    labels = list(top5.index) + ['Demais Veículos']
    valores = list(top5.values) + [outros_valor]
    
    # Calcula porcentagens sobre o total geral
    pcts = [(v / total_geral * 100) if total_geral > 0 else 0 for v in valores]
    
    # PALETA SPRES - 6 cores oficiais
    cores_spres = [
        SPRES_BLUE_DARK,      # Azul escuro
        SPRES_BLUE,           # Azul corporativo
        SPRES_BLUE_LIGHT,     # Azul claro
        SPRES_YELLOW_DARK,    # Amarelo escuro
        SPRES_YELLOW,         # Amarelo vivo
        SPRES_YELLOW_LIGHT   # amarelo leve
    ]
    
    # Cria o treemap - CORRIGIDO para evitar "undefined"
    fig = go.Figure(go.Treemap(
        labels=labels,
        parents=[''] * len(labels),  # Todos no mesmo nível (raiz)
        values=valores,
        
        # Template do texto dentro dos retângulos
        texttemplate='%{label}<br>R$ %{value:,.0f}<br>%{percentParent:.1%}',
        textfont=dict(
            size=13, 
            color='white', 
            family='Inter, sans-serif'
        ),
        
        # Cores Spres
        marker=dict(
            colors=cores_spres[:len(labels)],
            line=dict(color='rgba(255,255,255,0.4)', width=2.5),
            cornerradius=8
        ),
        
        # Tooltip enriquecido
        hovertemplate=(
            '<b>%{label}</b><br>' +
            'Valor: <b>R$ %{value:,.2f}</b><br>' +
            '% do Total: <b>%{percentRoot:.1%}</b><br>' +
            '<extra></extra>'
        ),
        
        textposition='middle center',
        tiling=dict(packing='squarify', pad=6),
        pathbar=dict(visible=False)
    ))
    
    # Layout limpo e profissional
    fig.update_layout(
        height=460,
        margin=dict(l=10, r=10, t=30, b=60),
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(color=SPRES_TEXT, family='Inter, sans-serif'),
        annotations=[
            dict(
                text=f'<i>Total: {formatar_moeda(total_geral)} • 6 principais categorias</i>',
                x=0.5,
                y=-0.05,
                xref='paper',
                yref='paper',
                showarrow=False,
                font=dict(size=11, color=SPRES_TEXT_MUTED)
            )
        ]
    )
    
    return fig

def criar_legenda_distribuicao(distribuicao_veiculo):
    """Cria HTML com lista dos veículos restantes (após top 8) com suas porcentagens"""

    total = distribuicao_veiculo.sum()
    resto = distribuicao_veiculo.sort_values(ascending=False).iloc[10:]

    if resto.empty:
        return None

    rows_html = ''
    for i, (veiculo, valor) in enumerate(resto.items()):
        pct = (valor / total * 100) if total > 0 else 0
        color = '#7B8794' if i % 2 == 0 else '#8A96A3'
        rows_html += '<tr>'
        rows_html += '<td style="padding:6px 10px; font-size:12px; color:' + color + '; border-bottom:1px solid rgba(0,75,141,0.06);">'
        rows_html += '<span style="display:inline-block; width:8px; height:8px; border-radius:2px; background:' + color + '; margin-right:8px;"></span>'
        rows_html += str(veiculo) + '</td>'
        rows_html += '<td style="padding:6px 10px; font-size:12px; color:' + SPRES_TEXT_MUTED + '; text-align:right; border-bottom:1px solid rgba(0,75,141,0.06);">'
        rows_html += formatar_moeda(valor) + '</td>'
        rows_html += '<td style="padding:6px 10px; font-size:12px; color:' + SPRES_TEXT_MUTED + '; text-align:right; border-bottom:1px solid rgba(0,75,141,0.06);">'
        rows_html += str(round(pct, 1)) + '%</td></tr>'

        html = '<div style="margin-top:12px; padding:14px 16px; background:rgba(255,255,255,0.5); border-radius:12px; border:1px solid ' + TECH_GLASS_BORDER + ';">'
        html += '<div style="font-size:12px; font-weight:700; color:' + SPRES_TEXT_MUTED + '; text-transform:uppercase; letter-spacing:0.6px; margin-bottom:8px;">'
        html += 'Demais veículos (' + str(len(resto)) + ' itens)</div>'
        html += '<table style="width:100%; border-collapse:collapse;">'
        html += '<thead><tr style="border-bottom:2px solid rgba(0,75,141,0.1);">'
        html += '<th style="padding:6px 10px; font-size:11px; color:' + SPRES_TEXT_MUTED + '; text-align:left; text-transform:uppercase; letter-spacing:0.4px;">Veículo</th>'
        html += '<th style="padding:6px 10px; font-size:11px; color:' + SPRES_TEXT_MUTED + '; text-align:right; text-transform:uppercase; letter-spacing:0.4px;">Valor</th>'
        html += '<th style="padding:6px 10px; font-size:11px; color:' + SPRES_TEXT_MUTED + '; text-align:right; text-transform:uppercase; letter-spacing:0.4px;">%</th>'
        html += '</tr></thead><tbody>'
        html += rows_html
        html += '</tbody></table></div>'

    return html

def formatar_celula_moeda(valor):
    """Formata célula: 0 vira '-' e demais valores viram moeda BR."""
    if pd.isna(valor) or valor == 0:
        return '-'
    return formatar_moeda(valor)


def criar_tabela_controle(dados_controle, big_numbers, totais_aprovado_mensal,
                          totais_utilizado_mensal, totais_saldo_mensal):
    """
    Cria tabela formatada do controle.
    NÃO adiciona linhas de total se elas já existirem no DataFrame.
    """
    
    # Colunas que queremos exibir
    colunas_exibir = ['veiculo', 'canal', 'obs', 'fev/26', 'mar/26', 'abr/26', 'mai/26',
                      'jun/26', 'jul/26', 'ago/26', 'set/26', 'out/26', 'nov/26', 'dez/26', 'jan/27', 'total']
    
    # Verifica quais colunas existem no DataFrame
    colunas_existentes = [col for col in colunas_exibir if col in dados_controle.columns]
    
    # Cria uma cópia apenas com as colunas existentes
    tabela = dados_controle[colunas_existentes].copy()
    
    # Renomeia colunas para exibição
    mapping = {
        'veiculo': 'Veículo', 'canal': 'Canal', 'obs': 'Observação',
        'fev/26': 'Fev/26', 'mar/26': 'Mar/26', 'abr/26': 'Abr/26',
        'mai/26': 'Mai/26', 'jun/26': 'Jun/26', 'jul/26': 'Jul/26',
        'ago/26': 'Ago/26', 'set/26': 'Set/26', 'out/26': 'Out/26',
        'nov/26': 'Nov/26', 'dez/26': 'Dez/26', 'jan/27': 'Jan/27',
        'total': 'Total'
    }
    tabela = tabela.rename(columns=mapping)
    
    # ============================================================
    # VERIFICA SE AS LINHAS DE TOTAL JÁ EXISTEM NO DATAFRAME
    # ============================================================
    # Procura por linhas que contenham "TOTAL" no Veículo
    if 'Veículo' in tabela.columns:
        # Verifica se já existem linhas de total
        linhas_total = tabela[tabela['Veículo'].str.contains('TOTAL', case=False, na=False)]
        
        # Se NÃO existir linhas de total, ADICIONA
        if linhas_total.empty:
            # Define os meses para as linhas de total
            colunas_mes = ['Fev/26', 'Mar/26', 'Abr/26', 'Mai/26', 'Jun/26', 'Jul/26',
                           'Ago/26', 'Set/26', 'Out/26', 'Nov/26', 'Dez/26', 'Jan/27']
            
            # Mapeia meses_short para meses_dict
            meses_dict = {
                'Fev/26': 'fev/26', 'Mar/26': 'mar/26', 'Abr/26': 'abr/26',
                'Mai/26': 'mai/26', 'Jun/26': 'jun/26', 'Jul/26': 'jul/26',
                'Ago/26': 'ago/26', 'Set/26': 'set/26', 'Out/26': 'out/26',
                'Nov/26': 'nov/26', 'Dez/26': 'dez/26', 'Jan/27': 'jan/27'
            }
            
            # Linha de Total Controle Aprovado
            total_aprovado_dict = {
                'Veículo': '📌 TOTAL CONTROLE APROVADO',
                'Canal': '',
                'Observação': 'Linha 47',
            }
            for mes in colunas_mes:
                mes_key = meses_dict.get(mes, mes.lower())
                total_aprovado_dict[mes] = totais_aprovado_mensal.get(mes_key, 0)
            total_aprovado_dict['Total'] = big_numbers.get('total_controle_aprovado', 0)
            
            # Linha de Total Controle Utilizado
            total_utilizado_dict = {
                'Veículo': '🍊 TOTAL CONTROLE UTILIZADO',
                'Canal': '',
                'Observação': 'Linha 48',
            }
            for mes in colunas_mes:
                mes_key = meses_dict.get(mes, mes.lower())
                total_utilizado_dict[mes] = totais_utilizado_mensal.get(mes_key, 0)
            total_utilizado_dict['Total'] = big_numbers.get('total_controle_utilizado', 0)
            
            # Linha de Total Saldo Positivo
            total_saldo_dict = {
                'Veículo': '💰 TOTAL SALDO POSITIVO',
                'Canal': '',
                'Observação': 'Linha 49',
            }
            for mes in colunas_mes:
                mes_key = meses_dict.get(mes, mes.lower())
                total_saldo_dict[mes] = totais_saldo_mensal.get(mes_key, 0)
            total_saldo_dict['Total'] = big_numbers.get('total_saldo_positivo', 0)
            
            # Adiciona as linhas de total
            total_row = pd.DataFrame([total_aprovado_dict])
            utilizado_row = pd.DataFrame([total_utilizado_dict])
            saldo_row = pd.DataFrame([total_saldo_dict])
            
            tabela = pd.concat([tabela, total_row, utilizado_row, saldo_row], ignore_index=True)
        else:
            # Se já existir linhas de total, não adiciona novamente
            # Apenas atualiza os valores se necessário
            for idx in linhas_total.index:
                if 'APROVADO' in str(linhas_total.loc[idx, 'Veículo']).upper():
                    tabela.loc[idx, 'Total'] = big_numbers.get('total_controle_aprovado', 0)
                elif 'UTILIZADO' in str(linhas_total.loc[idx, 'Veículo']).upper():
                    tabela.loc[idx, 'Total'] = big_numbers.get('total_controle_utilizado', 0)
                elif 'SALDO' in str(linhas_total.loc[idx, 'Veículo']).upper():
                    tabela.loc[idx, 'Total'] = big_numbers.get('total_saldo_positivo', 0)
    
    # ============================================================
    # FORMATAÇÃO DOS VALORES
    # ============================================================
    colunas_mes = ['Fev/26', 'Mar/26', 'Abr/26', 'Mai/26', 'Jun/26', 'Jul/26',
                   'Ago/26', 'Set/26', 'Out/26', 'Nov/26', 'Dez/26', 'Jan/27', 'Total']
    
    for col in colunas_mes:
        if col in tabela.columns:
            tabela[col] = tabela[col].apply(formatar_celula_moeda)
    
    return tabela


meses_short = ['Fev/26', 'Mar/26', 'Abr/26', 'Mai/26', 'Jun/26', 'Jul/26',
               'Ago/26', 'Set/26', 'Out/26', 'Nov/26', 'Dez/26', 'Jan/27']
meses_dict = {
    'Fev/26': 'fev/26', 'Mar/26': 'mar/26', 'Abr/26': 'abr/26',
    'Mai/26': 'mai/26', 'Jun/26': 'jun/26', 'Jul/26': 'jul/26',
    'Ago/26': 'ago/26', 'Set/26': 'set/26', 'Out/26': 'out/26',
    'Nov/26': 'nov/26', 'Dez/26': 'dez/26', 'Jan/27': 'jan/27'
}

# ============================================
# 4. FUNÇÃO DE EXPORTAÇÃO
# ============================================
def exportar_relatorio(dados_controle):
    """Exporta dados para CSV"""

    csv = dados_controle.to_csv(index=False, sep=';', decimal=',')
    b64 = base64.b64encode(csv.encode()).decode()
    href = f'<a href="data:file/csv;base64,{b64}" download="relatorio_midia_{datetime.now().strftime("%Y%m%d")}.csv" class="tech-btn-download">📥 Baixar Relatório CSV</a>'
    return href

# ============================================
# 5. LOGO SPRES (imagem oficial)
# ============================================
def logo_spres_base64():
    """Converte o logo real (logo.png) para base64, pronto para uso em HTML."""
    try:
        with open("logo.png", "rb") as img_file:
            return base64.b64encode(img_file.read()).decode()
    except Exception:
        return None

# ============================================
# 5.1. FUNÇÕES DE STATUS DO SHAREPOINT
# ============================================
def verificar_status_sharepoint():
    """Verifica e retorna status da conexão SharePoint"""
    if not SHAREPOINT_AVAILABLE:
        return False, "Módulo não disponível"
    
    if not GraphConfig.is_configured():
        return False, "Configuração incompleta"
    
    try:
        success, message = test_sharepoint_connection()
        return success, message
    except Exception as e:
        return False, f"Erro: {str(e)}"

def mostrar_status_sharepoint():
    """Exibe o status do SharePoint no sidebar"""
    if not SHAREPOINT_AVAILABLE:
        st.sidebar.warning("⚠️ Módulo SharePoint não disponível")
        st.sidebar.info("📌 Usando dados estáticos")
        return
    
    if not GraphConfig.is_configured():
        st.sidebar.warning("⚠️ SharePoint não configurado")
        st.sidebar.info("📌 Configure .env com credenciais")
        st.sidebar.code("""
AZURE_TENANT_ID=...
AZURE_CLIENT_ID=...
AZURE_CLIENT_SECRET=...
SHAREPOINT_SITE_URL=...
        """)
        return
    
    success, message = verificar_status_sharepoint()
    
    if success:
        st.sidebar.success("✅ SharePoint Online Conectado")
        st.sidebar.info(f"📁 {GraphConfig.EXCEL_FILENAME}")
        st.sidebar.caption(f"🔄 Cache: {GraphConfig.CACHE_MINUTES} min")
        
        if st.sidebar.button("🔄 Recarregar Dados", use_container_width=True, key="btn_recarregar_sidebar_status"):
            try:
                invalidate_sharepoint_cache()
                st.sidebar.success("✅ Cache invalidado! Recarregando...")
                time.sleep(1)
                st.rerun()
            except:
                st.sidebar.error("❌ Erro ao recarregar")
    else:
        st.sidebar.error(f"❌ {message}")
        st.sidebar.info("📌 Usando dados estáticos como fallback")

# ============================================
# 5.2. FUNÇÃO DE STATUS INTEGRADO
# ============================================
def mostrar_status_integracao():
    """Exibe o status da integração no sidebar"""
    if not EXCEL_INTEGRATION_AVAILABLE:
        st.sidebar.warning("⚠️ Módulo de integração Excel não disponível")
        return
    
    try:
        integracao = ExcelIntegration()
        status = integracao.get_status()
        
        # Status SharePoint
        if status.get('sharepoint_disponivel', False) and status.get('sharepoint_configurado', False):
            try:
                success, message = test_sharepoint_connection(quick=True)
                if success:
                    st.sidebar.success("✅ SharePoint Online Conectado")
                    st.sidebar.caption(f"📁 {GraphConfig.EXCEL_FILENAME}")
                    st.sidebar.caption(f"🔄 Cache: {GraphConfig.CACHE_MINUTES} min")
                else:
                    st.sidebar.warning(f"⚠️ {message}")
            except Exception as e:
                st.sidebar.warning(f"⚠️ SharePoint: Erro na conexão - {str(e)}")
        else:
            st.sidebar.info("📌 SharePoint não configurado")
            st.sidebar.caption("Configure o .env com as credenciais")
        
        # Última atualização
        st.sidebar.caption(f"🔄 Última atualização: {status.get('ultima_atualizacao', 'Nunca')}")
        st.sidebar.caption(f"📁 Origem: {status.get('origem', 'Não carregado')}")
        
        # Quantidade de dados
        dados_controle = status.get('dados_controle', 0)
        if dados_controle > 0:
            st.sidebar.caption(f"📊 Itens carregados: {dados_controle}")
        
        # Botão de recarga
        if st.sidebar.button("🔄 Recarregar Dados", use_container_width=True, key="btn_recarregar_integracao"):
            try:
                integracao.forcar_atualizacao()
                if integracao.carregar_sharepoint(forcar=True):
                    st.sidebar.success("✅ Dados recarregados com sucesso!")
                else:
                    st.sidebar.warning("⚠️ Falha ao recarregar. Usando dados em cache.")
                time.sleep(0.5)
                st.rerun()
            except Exception as e:
                st.sidebar.error(f"❌ Erro: {str(e)}")
                
    except Exception as e:
        st.sidebar.error(f"❌ Erro no status: {str(e)}")

# ============================================
# 6. INTERFACE PRINCIPAL — LAYOUT TECNOLÓGICO
# ============================================
def main():
    # ============================================================
    # LIMPA CONTEÚDO DA TELA DE LOGIN E CONFIGURA SIDEBAR
    # ============================================================
    st.markdown("""
    <style>
        /* ===== REMOVE CONTEÚDO DA TELA DE LOGIN ===== */
        .stApp {
            background: none !important;
            animation: none !important;
        }
        .login-particles {
            display: none !important;
        }
        .login-card-wrapper {
            display: none !important;
        }
        .login-card-inner {
            display: none !important;
        }
        .login-logo, .login-subtitle, .login-title {
            display: none !important;
        }
        div[data-testid="stForm"] {
            display: none !important;
        }
        .login-error-msg, .login-attempts, .login-footer {
            display: none !important;
        }
        .block-container {
            padding-top: 0 !important;
        }

        /* ===== SIDEBAR SPRES ===== */
        section[data-testid="stSidebar"] {
            background: linear-gradient(180deg, #0A1F35 0%, #00325F 50%, #004B8D 100%) !important;
            border-right: 2px solid rgba(255, 214, 0, 0.15) !important;
        }
        
        section[data-testid="stSidebar"] .stMarkdown {
            color: #FFFFFF !important;
        }
        
        section[data-testid="stSidebar"] .stMarkdown h1,
        section[data-testid="stSidebar"] .stMarkdown h2,
        section[data-testid="stSidebar"] .stMarkdown h3 {
            color: #FFFFFF !important;
        }
        
        section[data-testid="stSidebar"] .stMarkdown p,
        section[data-testid="stSidebar"] .stMarkdown li,
        section[data-testid="stSidebar"] .stMarkdown div {
            color: rgba(255, 255, 255, 0.85) !important;
        }
        
        section[data-testid="stSidebar"] .stMarkdown .st-caption {
            color: rgba(255, 255, 255, 0.5) !important;
        }
        
        section[data-testid="stSidebar"] .stMarkdown code {
            background: rgba(255, 255, 255, 0.08) !important;
            color: #FFD600 !important;
        }

        /* ===== BOTÕES DO SIDEBAR ===== */
        section[data-testid="stSidebar"] .stButton button {
            background: linear-gradient(135deg, #FFD600 0%, #FFB300 100%) !important;
            color: #00325F !important;
            border: none !important;
            font-weight: 700 !important;
            transition: all 0.3s ease !important;
            box-shadow: 0 2px 8px rgba(255, 214, 0, 0.2) !important;
        }
        
        section[data-testid="stSidebar"] .stButton button:hover {
            transform: translateY(-2px) !important;
            box-shadow: 0 4px 16px rgba(255, 214, 0, 0.4) !important;
        }
        
        section[data-testid="stSidebar"] .stButton button:active {
            transform: translateY(0) !important;
        }

        /* ===== BOTÃO DE UPLOAD (CORRIGIDO) ===== */
        section[data-testid="stSidebar"] .stFileUploader button {
            background: rgba(255, 214, 0, 0.2) !important;
            color: #FFD600 !important;
            border: 1px solid rgba(255, 214, 0, 0.3) !important;
            font-weight: 600 !important;
            transition: all 0.3s ease !important;
        }
        
        section[data-testid="stSidebar"] .stFileUploader button:hover {
            background: rgba(255, 214, 0, 0.3) !important;
            border-color: #FFD600 !important;
        }
        
        /* Texto do upload */
        section[data-testid="stSidebar"] .stFileUploader label {
            color: rgba(255, 255, 255, 0.8) !important;
        }
        
        section[data-testid="stSidebar"] .stFileUploader .st-dz-message {
            color: rgba(255, 255, 255, 0.6) !important;
        }

        /* ===== BOTÃO "PROCESSAR PLANILHA" ===== */
        section[data-testid="stSidebar"] .stButton button[data-testid="baseButton-secondary"] {
            background: rgba(255, 214, 0, 0.15) !important;
            color: #FFD600 !important;
            border: 1px solid rgba(255, 214, 0, 0.2) !important;
        }
        
        section[data-testid="stSidebar"] .stButton button[data-testid="baseButton-secondary"]:hover {
            background: rgba(255, 214, 0, 0.25) !important;
            border-color: #FFD600 !important;
        }

        /* ===== EXPANDER ===== */
        section[data-testid="stSidebar"] .streamlit-expanderHeader {
            color: rgba(255, 255, 255, 0.9) !important;
            background: rgba(255, 255, 255, 0.05) !important;
            border-radius: 8px !important;
        }
        
        section[data-testid="stSidebar"] .streamlit-expanderContent {
            background: rgba(255, 255, 255, 0.03) !important;
            border-radius: 0 0 8px 8px !important;
        }

        /* ===== SEPARADORES ===== */
        section[data-testid="stSidebar"] hr {
            border-color: rgba(255, 214, 0, 0.15) !important;
        }

        /* ===== ALERTAS ===== */
        section[data-testid="stSidebar"] .stAlert {
            background: rgba(255, 214, 0, 0.08) !important;
            border-color: rgba(255, 214, 0, 0.15) !important;
            color: rgba(255, 255, 255, 0.9) !important;
        }
        
        section[data-testid="stSidebar"] .stAlert svg {
            color: #FFD600 !important;
        }

        /* ===== MÉTRICAS ===== */
        section[data-testid="stSidebar"] div[data-testid="stMetric"] {
            background: rgba(255, 255, 255, 0.05) !important;
            border: 1px solid rgba(255, 255, 255, 0.08) !important;
            border-radius: 12px !important;
            padding: 12px !important;
        }
        
        section[data-testid="stSidebar"] div[data-testid="stMetric"] label {
            color: rgba(255, 255, 255, 0.6) !important;
        }
        
        section[data-testid="stSidebar"] div[data-testid="stMetric"] div {
            color: #FFD600 !important;
        }

        /* ===== STATUS BADGE ===== */
        .tech-status-badge {
            background: rgba(255, 214, 0, 0.12) !important;
            border: 1px solid rgba(255, 214, 0, 0.15) !important;
            color: #FFD600 !important;
        }
        
        .tech-status-active {
            background: rgba(255, 214, 0, 0.12) !important;
            color: #FFD600 !important;
            border: 1px solid rgba(255, 214, 0, 0.15) !important;
        }

        /* ===== INPUTS NO SIDEBAR ===== */
        section[data-testid="stSidebar"] .stTextInput input {
            background: rgba(255, 255, 255, 0.05) !important;
            color: #FFFFFF !important;
            border: 1px solid rgba(255, 255, 255, 0.1) !important;
            border-radius: 8px !important;
        }
        
        section[data-testid="stSidebar"] .stTextInput input:focus {
            border-color: #FFD600 !important;
            box-shadow: 0 0 0 2px rgba(255, 214, 0, 0.1) !important;
        }
    </style>
    """, unsafe_allow_html=True)
    
    # Limpa qualquer conteúdo residual
    for _ in range(2):
        st.empty()
    
    # ===== VERIFICA SE ESTÁ AUTENTICADO =====
    if not st.session_state.get("autenticado", False):
        pagina_login()
        return
    
    # ===== CSS DO DASHBOARD =====
    st.markdown(f"""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap');

        /* ========== BASE ========== */
        .stApp {{
            background: linear-gradient(135deg, #F7FAFC 0%, #EDF4FB 40%, #F0F7FF 100%);
            font-family: 'Inter', 'Segoe UI', sans-serif;
            min-height: 100vh;
        }}

        /* Background grid pattern (sutil) */
        .stApp::before {{
            content: '';
            position: fixed;
            top: 0; left: 0; right: 0; bottom: 0;
            background-image:
                radial-gradient(circle at 20% 50%, rgba(0,75,141,0.03) 0%, transparent 50%),
                radial-gradient(circle at 80% 20%, rgba(255,214,0,0.04) 0%, transparent 50%),
                radial-gradient(circle at 50% 80%, rgba(0,229,255,0.02) 0%, transparent 50%);
            pointer-events: none;
            z-index: 0;
        }}

        /* ========== HEADER TECNOLÓGICO ========== */
        .main-header {{
            background: linear-gradient(135deg, {SPRES_BLUE_DARK} 0%, {SPRES_BLUE} 60%, {SPRES_BLUE_LIGHT} 100%);
            padding: 22px 32px;
            border-radius: 20px;
            box-shadow:
                0 8px 32px rgba(0, 50, 95, 0.35),
                0 0 0 1px rgba(255, 214, 0, 0.2),
                inset 0 1px 0 rgba(255, 255, 255, 0.1);
            margin-bottom: 28px;
            display: flex;
            align-items: center;
            gap: 22px;
            position: relative;
            overflow: hidden;
        }}
        .main-header::before {{
            content: '';
            position: absolute;
            top: -50%;
            right: -10%;
            width: 300px;
            height: 300px;
            background: radial-gradient(circle, rgba(255,214,0,0.08) 0%, transparent 70%);
            border-radius: 50%;
            pointer-events: none;
        }}
        .main-header::after {{
            content: '';
            position: absolute;
            bottom: 0;
            left: 0;
            right: 0;
            height: 3px;
            background: linear-gradient(90deg, {SPRES_YELLOW}, {SPRES_ORANGE}, {SPRES_YELLOW});
            border-radius: 0 0 20px 20px;
        }}
        .main-header img {{
            filter: drop-shadow(0 4px 8px rgba(0,0,0,0.3));
            position: relative;
            z-index: 1;
        }}
        .main-header h1 {{
            color: {SPRES_WHITE};
            font-size: 30px;
            font-weight: 800;
            margin: 0;
            letter-spacing: -0.5px;
            position: relative;
            z-index: 1;
        }}
        .main-header span {{
            color: {SPRES_YELLOW};
            text-shadow: 0 0 20px rgba(255, 214, 0, 0.4);
        }}
        .sub-header {{
            color: #b8d4f0;
            font-size: 14px;
            margin-top: 4px;
            font-weight: 400;
            letter-spacing: 0.3px;
            position: relative;
            z-index: 1;
        }}

        /* ========== TECH CARDS (Glassmorphism) ========== */
        .card-container {{
            background: {TECH_GLASS_BG};
            backdrop-filter: blur(20px);
            -webkit-backdrop-filter: blur(20px);
            border-radius: 18px;
            padding: 24px;
            box-shadow:
                0 4px 24px rgba(0, 75, 141, 0.06),
                0 0 0 1px rgba(0, 75, 141, 0.08);
            border: 1px solid {TECH_GLASS_BORDER};
            margin-bottom: 22px;
            position: relative;
            overflow: hidden;
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        }}
        .card-container:hover {{
            box-shadow:
                0 8px 32px rgba(0, 75, 141, 0.10),
                0 0 0 1px rgba(0, 75, 141, 0.14);
            transform: translateY(-2px);
        }}
        .card-container::before {{
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            height: 2px;
            background: linear-gradient(90deg, transparent, {SPRES_BLUE_LIGHT}, {SPRES_YELLOW}, {SPRES_BLUE_LIGHT}, transparent);
            opacity: 0.6;
        }}
        .card-title {{
            font-size: 16px;
            font-weight: 700;
            color: {SPRES_BLUE_DARK};
            margin-bottom: 18px;
            display: flex;
            align-items: center;
            gap: 10px;
            letter-spacing: -0.3px;
        }}
        .card-title::before {{
            content: '';
            display: inline-block;
            width: 4px;
            height: 18px;
            background: linear-gradient(180deg, {SPRES_BLUE}, {SPRES_BLUE_LIGHT});
            border-radius: 2px;
            flex-shrink: 0;
        }}

        /* ========== BIG NUMBERS TECNOLÓGICOS ========== */
        .big-number-container {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
            gap: 18px;
            margin: 0 0 26px 0;
        }}
        .big-number-item {{
            background: {TECH_GLASS_BG};
            backdrop-filter: blur(16px);
            -webkit-backdrop-filter: blur(16px);
            border-radius: 18px;
            padding: 22px 20px;
            border: 1px solid {TECH_GLASS_BORDER};
            position: relative;
            overflow: hidden;
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        }}
        .big-number-item:hover {{
            transform: translateY(-4px);
            box-shadow: 0 12px 36px rgba(0, 75, 141, 0.15);
        }}

        .tech-glow-blue {{
            border-top: 3px solid {SPRES_BLUE};
            box-shadow: 0 4px 16px rgba(0, 75, 141, 0.08);
        }}
        .tech-glow-blue::after {{
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            height: 40px;
            background: linear-gradient(180deg, rgba(0,75,141,0.06) 0%, transparent 100%);
            pointer-events: none;
        }}
        .tech-glow-yellow {{
            border-top: 3px solid {SPRES_YELLOW};
            box-shadow: 0 4px 16px rgba(255, 214, 0, 0.10);
        }}
        .tech-glow-yellow::after {{
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            height: 40px;
            background: linear-gradient(180deg, rgba(255,214,0,0.08) 0%, transparent 100%);
            pointer-events: none;
        }}
        .tech-glow-orange {{
            border-top: 3px solid {SPRES_ORANGE};
            box-shadow: 0 4px 16px rgba(255, 138, 30, 0.08);
        }}
        .tech-glow-orange::after {{
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            height: 40px;
            background: linear-gradient(180deg, rgba(255,138,30,0.06) 0%, transparent 100%);
            pointer-events: none;
        }}

        .big-number-header {{
            display: flex;
            align-items: center;
            gap: 10px;
            margin-bottom: 10px;
        }}
        .big-number-icon-tech {{
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 36px;
            height: 36px;
            border-radius: 10px;
            font-size: 16px;
            font-weight: 700;
            position: relative;
            z-index: 1;
        }}
        .tech-glow-blue .big-number-icon-tech {{
            background: linear-gradient(135deg, rgba(0,75,141,0.10), rgba(0,75,141,0.04));
            color: {SPRES_BLUE};
        }}
        .tech-glow-yellow .big-number-icon-tech {{
            background: linear-gradient(135deg, rgba(255,214,0,0.15), rgba(255,214,0,0.05));
            color: {SPRES_YELLOW_DARK};
        }}
        .tech-glow-orange .big-number-icon-tech {{
            background: linear-gradient(135deg, rgba(255,138,30,0.12), rgba(255,138,30,0.04));
            color: {SPRES_ORANGE};
        }}
        .big-number-pct {{
            font-size: 22px;
            font-weight: 800;
            color: {SPRES_TEXT_MUTED};
            margin-left: auto;
        }}
        .big-number-label {{
            font-size: 12px;
            color: {SPRES_TEXT_MUTED};
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.8px;
            margin-bottom: 8px;
        }}
        .big-number-value {{
            font-size: 28px;
            font-weight: 800;
            color: {SPRES_BLUE_DARK};
            letter-spacing: -0.5px;
        }}
        .big-number-sub {{
            font-size: 12px;
            color: {SPRES_TEXT_MUTED};
            margin-top: 6px;
            font-weight: 400;
        }}
        .big-number-bar {{
            margin-top: 12px;
            height: 6px;
            background: rgba(0,75,141,0.06);
            border-radius: 3px;
            overflow: hidden;
        }}
        .big-number-bar-fill {{
            height: 100%;
            border-radius: 3px;
            transition: width 0.6s ease;
        }}

        /* ========== stMetric customizado ========== */
        div[data-testid="stMetric"] {{
            background: {TECH_GLASS_BG};
            backdrop-filter: blur(12px);
            -webkit-backdrop-filter: blur(12px);
            border-radius: 16px;
            border: 1px solid {TECH_GLASS_BORDER};
            border-left: 4px solid {SPRES_BLUE};
            padding: 18px;
            box-shadow: 0 2px 12px rgba(0, 75, 141, 0.05);
            transition: all 0.2s ease;
        }}
        div[data-testid="stMetric"]:hover {{
            box-shadow: 0 4px 20px rgba(0, 75, 141, 0.10);
            transform: translateY(-1px);
        }}
        div[data-testid="stMetric"] label p {{
            color: {SPRES_TEXT_MUTED} !important;
            font-weight: 600 !important;
            font-size: 12px !important;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }}
        div[data-testid="stMetric"] div[data-testid="stMetricValue"] {{
            color: {SPRES_BLUE_DARK} !important;
            font-weight: 800 !important;
            font-size: 22px !important;
            letter-spacing: -0.3px;
        }}

        /* ========== Abas (tabs) tecnológicas ========== */
        .stTabs [data-baseweb="tab-list"] {{
            gap: 8px;
            margin-bottom: 20px;
            padding: 8px;
            background: rgba(255, 255, 255, 0.5);
            border-radius: 16px;
            border: 1px solid {TECH_GLASS_BORDER};
        }}
        .stTabs [data-baseweb="tab"] {{
            background-color: transparent;
            border-radius: 12px;
            color: {SPRES_TEXT_MUTED};
            font-weight: 600;
            padding: 14px 28px;
            border: none;
            font-size: 14px;
            min-width: 160px;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
            transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1);
        }}
        .stTabs [data-baseweb="tab"]:hover {{
            background-color: rgba(0, 75, 141, 0.04) !important;
            color: {SPRES_BLUE} !important;
        }}
        .stTabs [aria-selected="true"] {{
            background: linear-gradient(135deg, {SPRES_BLUE} 0%, {SPRES_BLUE_DARK} 100%) !important;
            color: {SPRES_WHITE} !important;
            box-shadow: 0 4px 16px rgba(0, 75, 141, 0.3);
            padding: 14px 28px !important;
            min-width: 160px;
        }}

        /* ========== Dataframes tecnológicos ========== */
        [data-testid="stDataFrame"] {{
            border-radius: 16px !important;
            overflow: hidden !important;
            border: 1px solid {TECH_GLASS_BORDER} !important;
            box-shadow: 0 2px 12px rgba(0, 75, 141, 0.04) !important;
        }}

        /* ========== Botões / links ========== */
        .tech-btn-download {{
            display: inline-flex;
            align-items: center;
            gap: 8px;
            margin-top: 16px;
            padding: 12px 24px;
            background: linear-gradient(135deg, {SPRES_BLUE}, {SPRES_BLUE_DARK});
            color: {SPRES_WHITE} !important;
            border-radius: 12px;
            text-decoration: none;
            font-weight: 600;
            font-size: 14px;
            transition: all 0.2s ease;
            box-shadow: 0 4px 16px rgba(0, 75, 141, 0.25);
            letter-spacing: 0.2px;
        }}
        .tech-btn-download:hover {{
            background: linear-gradient(135deg, {SPRES_BLUE_DARK}, {SPRES_BLUE});
            transform: translateY(-2px);
            box-shadow: 0 6px 24px rgba(0, 75, 141, 0.35);
        }}
        
        /* ========== Footer ========== */
        .footer {{
            text-align: center;
            color: {SPRES_TEXT_MUTED};
            font-size: 13px;
            margin-top: 36px;
            padding-top: 20px;
            border-top: 1px solid {TECH_GLASS_BORDER};
            font-weight: 400;
            letter-spacing: 0.2px;
        }}
        .footer strong {{
            color: {SPRES_BLUE};
        }}

        /* ========== Status badge ========== */
        .tech-status-badge {{
            display: inline-flex;
            align-items: center;
            gap: 6px;
            padding: 5px 12px;
            border-radius: 20px;
            font-size: 11px;
            font-weight: 600;
            letter-spacing: 0.3px;
        }}
        .tech-status-active {{
            background: rgba(0, 75, 141, 0.08);
            color: {SPRES_BLUE};
            border: 1px solid rgba(0, 75, 141, 0.15);
        }}
        .tech-status-sharepoint {{
            background: rgba(46, 125, 209, 0.10);
            color: {SPRES_BLUE_LIGHT};
            border: 1px solid rgba(46, 125, 209, 0.20);
        }}
        .tech-status-fallback {{
            background: rgba(255, 138, 30, 0.08);
            color: {SPRES_ORANGE};
            border: 1px solid rgba(255, 138, 30, 0.15);
        }}

        /* ========== Pulse animation ========== */
        @keyframes pulse-subtle {{
            0%, 100% {{ opacity: 1; }}
            50% {{ opacity: 0.7; }}
        }}
        .tech-pulse {{
            animation: pulse-subtle 3s ease-in-out infinite;
        }}
        @keyframes pulse-sharepoint {{
            0%, 100% {{ opacity: 1; }}
            50% {{ opacity: 0.5; }}
        }}
        .tech-pulse-sharepoint {{
            animation: pulse-sharepoint 2s ease-in-out infinite;
        }}
    </style>
    """, unsafe_allow_html=True)

    # ============================================================
    # CONTROLE DE CARREGAMENTO DE DADOS
    # ============================================================
    # Verifica se o usuário está logado e se os dados já foram carregados
    if st.session_state.get("autenticado", False):
        if not st.session_state.get("dados_carregados", False):
            # Mostra spinner enquanto carrega os dados
            with st.spinner("🔄 Carregando dados do SharePoint em tempo real..."):
                # Carrega os dados usando a lógica existente
                try:
                    # Usa a integração Excel se disponível
                    if EXCEL_INTEGRATION_AVAILABLE:
                        integracao = ExcelIntegration()
                        if SHAREPOINT_AVAILABLE and GraphConfig.is_configured():
                            integracao.carregar_sharepoint()
                        if integracao.dados_controle is not None:
                            (dados_geral, dados_controle, big_numbers, totais_aprovado_mensal,
                             totais_utilizado_mensal, totais_saldo_mensal, dados_estudo, dados_extra) = integracao.get_dados_atualizados()
                        else:
                            (dados_geral, dados_controle, big_numbers, totais_aprovado_mensal,
                             totais_utilizado_mensal, totais_saldo_mensal, dados_estudo, dados_extra) = carregar_dados_estaticos()
                    else:
                        # Fallback para carregamento tradicional
                        if SHAREPOINT_AVAILABLE and GraphConfig.is_configured():
                            (dados_geral, dados_controle, big_numbers, totais_aprovado_mensal,
                             totais_utilizado_mensal, totais_saldo_mensal, dados_estudo, dados_extra) = carregar_dados_sharepoint()
                        else:
                            (dados_geral, dados_controle, big_numbers, totais_aprovado_mensal,
                             totais_utilizado_mensal, totais_saldo_mensal, dados_estudo, dados_extra) = carregar_dados_estaticos()
                    
                    # Salva no session_state para uso posterior
                    st.session_state.dados_geral = dados_geral
                    st.session_state.dados_controle = dados_controle
                    st.session_state.big_numbers = big_numbers
                    st.session_state.totais_aprovado_mensal = totais_aprovado_mensal
                    st.session_state.totais_utilizado_mensal = totais_utilizado_mensal
                    st.session_state.totais_saldo_mensal = totais_saldo_mensal
                    st.session_state.dados_estudo = dados_estudo
                    st.session_state.dados_extra = dados_extra
                    st.session_state.dados_carregados = True
                    
                    # ===== MENSAGEM DE SUCESSO COM AUTO-REMOÇÃO =====
                    placeholder = st.empty()
                    placeholder.success("✅ Dados carregados com sucesso!")
                    time.sleep(2)
                    placeholder.empty()
                    
                except Exception as e:
                    # ===== MENSAGEM DE ERRO =====
                    placeholder = st.empty()
                    placeholder.warning(f"⚠️ Erro ao carregar dados: {str(e)}")
                    time.sleep(3)
                    placeholder.empty()
                    
                    # Fallback para dados estáticos
                    (dados_geral, dados_controle, big_numbers, totais_aprovado_mensal,
                     totais_utilizado_mensal, totais_saldo_mensal, dados_estudo, dados_extra) = carregar_dados_estaticos()
                    st.session_state.dados_geral = dados_geral
                    st.session_state.dados_controle = dados_controle
                    st.session_state.big_numbers = big_numbers
                    st.session_state.totais_aprovado_mensal = totais_aprovado_mensal
                    st.session_state.totais_utilizado_mensal = totais_utilizado_mensal
                    st.session_state.totais_saldo_mensal = totais_saldo_mensal
                    st.session_state.dados_estudo = dados_estudo
                    st.session_state.dados_extra = dados_extra
                    st.session_state.dados_carregados = True

    # ============================================================
    # RECUPERA DADOS DO SESSION_STATE
    # ============================================================
    # Se os dados já foram carregados, usa do session_state
    if st.session_state.get("dados_carregados", False):
        dados_geral = st.session_state.get("dados_geral")
        dados_controle = st.session_state.get("dados_controle")
        big_numbers = st.session_state.get("big_numbers")
        totais_aprovado_mensal = st.session_state.get("totais_aprovado_mensal")
        totais_utilizado_mensal = st.session_state.get("totais_utilizado_mensal")
        totais_saldo_mensal = st.session_state.get("totais_saldo_mensal")
        dados_estudo = st.session_state.get("dados_estudo")
        dados_extra = st.session_state.get("dados_extra")
    else:
        # Fallback: carrega dados estáticos
        (dados_geral, dados_controle, big_numbers, totais_aprovado_mensal,
         totais_utilizado_mensal, totais_saldo_mensal, dados_estudo, dados_extra) = carregar_dados_estaticos()

    # ============================================================
    # SIDEBAR - COM TODAS AS FUNCIONALIDADES DE GERENCIAMENTO
    # ============================================================
    with st.sidebar:
        st.markdown("### 📊 Status da Integração")
        
        # Usa a nova função de status integrado
        if EXCEL_INTEGRATION_AVAILABLE:
            mostrar_status_integracao()
        else:
            mostrar_status_sharepoint()
        
        st.markdown("---")
        
        # # ===== UPLOAD DE PLANILHA =====
        # if EXCEL_INTEGRATION_AVAILABLE:
        #     with st.expander("📤 Upload de Planilha"):
        #         criar_interface_upload()
        #     st.markdown("---")
        
        # ===== BOTÃO PARA ABRIR PLANILHA NO SHAREPOINT =====
        st.markdown("### 📂 Acesso Rápido")
        
        def criar_botao_abrir_excel():
            """Cria botão/link para abrir o Excel no SharePoint"""
            
            excel_url = os.getenv("SHAREPOINT_FILE_URL", "")
            
            if not excel_url:
                sharepoint_url = os.getenv("SHAREPOINT_SITE_URL", "")
                excel_filename = os.getenv("EXCEL_FILENAME", "base_spres_projeto_refatorada.xlsx")
                
                if not sharepoint_url:
                    return None
                
                sharepoint_url = sharepoint_url.rstrip('/')
                
                if "my.sharepoint.com" in sharepoint_url:
                    excel_url = f"{sharepoint_url}/Documents/{excel_filename}"
                else:
                    excel_url = f"{sharepoint_url}/{excel_filename}"
            
            # HTML com cores corrigidas para melhor legibilidade
            html_link = f"""
            <div style="margin: 10px 0;">
                <a href="{excel_url}" target="_blank" style="
                    display: inline-flex;
                    align-items: center;
                    gap: 10px;
                    padding: 12px 20px;
                    background: linear-gradient(135deg, #004B8D 0%, #2E7DD1 100%);
                    color: #FFFFFF !important;
                    border-radius: 12px;
                    text-decoration: none;
                    font-weight: 700;
                    font-size: 14px;
                    width: 100%;
                    transition: all 0.3s ease;
                    box-shadow: 0 4px 16px rgba(0, 75, 141, 0.35);
                    border: 1px solid rgba(255, 255, 255, 0.15);
                    text-shadow: 0 1px 2px rgba(0, 0, 0, 0.2);
                ">
                    <span style="font-size: 20px;">📊</span>
                    <span style="color: #FFFFFF; font-weight: 600;">Abrir Planilha no SharePoint</span>
                    <span style="margin-left: auto; font-size: 16px; color: #FFFFFF; opacity: 0.8;">↗</span>
                </a>
                <div style="font-size: 11px; color: #5B6E80; margin-top: 4px; text-align: center;">
                    {excel_url.split('/')[-1] if '/' in excel_url else 'Arquivo'}
                </div>
            </div>
            """
            
            return html_link
        
        if SHAREPOINT_AVAILABLE and GraphConfig.is_configured():
            html_link = criar_botao_abrir_excel()
            if html_link:
                st.markdown(html_link, unsafe_allow_html=True)
        else:
            st.info("📌 SharePoint não configurado")
        
        st.markdown("---")
        
        # ===== BOTÃO RECARREGAR =====
        st.markdown("### ⚡ Ações Rápidas")
        if st.button("🔄 Recarregar Dados", use_container_width=True, key="btn_recarregar_sidebar"):
            try:
                if SHAREPOINT_AVAILABLE:
                    invalidate_sharepoint_cache()
                st.cache_data.clear()
                st.session_state.dados_carregados = False
                st.success("✅ Cache limpo! Recarregando...")
                time.sleep(0.5)
                st.rerun()
            except Exception as e:
                st.error(f"❌ Erro: {str(e)}")
        
        st.markdown("---")
        st.markdown("### 📅 Sessão")
        st.caption(f"Usuário: {st.session_state.get('usuario', 'guest')}")
        st.caption(f"Atualizado: {datetime.now().strftime('%H:%M:%S')}")
        
        # Botão de logout
        if st.button("🚪 Sair", use_container_width=True, key="btn_logout"):
            for key in ['autenticado', 'usuario', 'login_time']:
                if key in st.session_state:
                    del st.session_state[key]
            st.rerun()
            
    # ============================================================
    # HEADER
    # ============================================================
    logo_b64 = logo_spres_base64()
    logo_html = f'<img src="data:image/png;base64,{logo_b64}" width="78">' if logo_b64 else "🍊"

    sharepoint_status = "LIVE"
    status_class = "tech-status-active"
    if SHAREPOINT_AVAILABLE and GraphConfig.is_configured():
        success, _ = verificar_status_sharepoint()
        if success:
            sharepoint_status = "📡 SharePoint"
            status_class = "tech-status-sharepoint tech-pulse-sharepoint"
        else:
            sharepoint_status = "📊 Fallback"
            status_class = "tech-status-fallback"
    else:
        sharepoint_status = "📊 Fallback"
        status_class = "tech-status-fallback"

    st.markdown(f"""
    <div class="main-header">
        {logo_html}
        <div>
            <h1>Gestão <span>Mídia</span> Pro</h1>
            <div class="sub-header">Dashboard de cronograma de mídia • Sucos Spres • Ciclo fev/26 – jan/27</div>
        </div>
        <div style="margin-left: auto; display: flex; gap: 12px; align-items: center;">
            <span class="tech-status-badge {status_class}">
                ● {sharepoint_status}
            </span>
            <span class="tech-status-badge tech-status-active tech-pulse">
                ● LIVE
            </span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ============================================================
    # MENU DE NAVEGAÇÃO - APENAS 3 ABAS
    # ============================================================
    menu_tabs = st.tabs(["📊 Dashboard", "📋 Cronograma Completo", "📈 Análises"])

    # Processa os dados para visualizações
    dados_controle_processado, distribuicao_veiculo, dados_mensais = processar_dados(dados_controle)

    # ============================================================
    # ABA 1: DASHBOARD
    # ============================================================
    with menu_tabs[0]:  # Dashboard
        criar_big_numbers(big_numbers, totais_aprovado_mensal, totais_utilizado_mensal, totais_saldo_mensal)

        # ===== CARDS "MÉTRICAS PRINCIPAIS" REMOVIDOS =====
        # (Os Big Numbers já mostram as informações principais)

        # Linha 1: Treemap ocupando largura total (mais espaço)
        st.markdown('<div class="card-container">', unsafe_allow_html=True)
        st.markdown('<div class="card-title">📊 Distribuição por Veículo</div>', unsafe_allow_html=True)
        fig_dist = criar_grafico_distribuicao(distribuicao_veiculo)
        st.plotly_chart(fig_dist, use_container_width=True, config={'displayModeBar': False})
        legenda_html = criar_legenda_distribuicao(distribuicao_veiculo)
        if legenda_html is not None:
            st.markdown(legenda_html, unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

        # Linha 2: Investimento Mensal + Resumo por Veículo lado a lado
        col1, col2 = st.columns([1, 1])

        with col1:
            with st.container():
                st.markdown('<div class="chart-container card-container" style="margin-top:14px;">', unsafe_allow_html=True)
                st.markdown('<div class="card-title">📈 Investimento Mensal</div>', unsafe_allow_html=True)
                fig_mensal = criar_grafico_mensal(dados_mensais)
                st.plotly_chart(fig_mensal, use_container_width=True, config={'displayModeBar': False})
                st.markdown('</div>', unsafe_allow_html=True)

        with col2:
            with st.container():
                st.markdown('<div class="card-container" style="margin-top:14px;">', unsafe_allow_html=True)
                st.markdown('<div class="card-title">📋 Resumo por Veículo</div>', unsafe_allow_html=True)

                resumo_veiculo = dados_controle_processado.groupby('veiculo').agg({
                    'total': 'sum',
                    'fev/26': 'sum', 'mar/26': 'sum', 'abr/26': 'sum', 'mai/26': 'sum',
                    'jun/26': 'sum', 'jul/26': 'sum', 'ago/26': 'sum', 'set/26': 'sum',
                    'out/26': 'sum', 'nov/26': 'sum', 'dez/26': 'sum', 'jan/27': 'sum'
                }).round(2)

                resumo_veiculo = resumo_veiculo.rename(columns={
                    'total': 'Total', 'fev/26': 'Fev', 'mar/26': 'Mar', 'abr/26': 'Abr',
                    'mai/26': 'Mai', 'jun/26': 'Jun', 'jul/26': 'Jul', 'ago/26': 'Ago',
                    'set/26': 'Set', 'out/26': 'Out', 'nov/26': 'Nov', 'dez/26': 'Dez', 'jan/27': 'Jan'
                })

                for col in resumo_veiculo.columns:
                    resumo_veiculo[col] = resumo_veiculo[col].apply(lambda x: formatar_moeda(x) if x > 0 else '-')

                st.dataframe(resumo_veiculo, use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)

    # ============================================================
    # ABA 2: CRONOGRAMA COMPLETO
    # ============================================================
    with menu_tabs[1]:  # Cronograma Completo
        st.markdown('<div class="card-container">', unsafe_allow_html=True)
        st.markdown('<div class="card-title">📋 Cronograma Completo de Mídia</div>', unsafe_allow_html=True)
        st.caption("Base: aba 'Geral Controle (2) nova' com Total Controle Aprovado, Utilizado e Saldo Positivo.")

        tabela = criar_tabela_controle(dados_controle_processado, big_numbers,
                                        totais_aprovado_mensal, totais_utilizado_mensal, totais_saldo_mensal)
        st.dataframe(tabela, use_container_width=True, hide_index=True)

        st.markdown(exportar_relatorio(dados_controle_processado), unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

    # ============================================================
    # ABA 3: ANÁLISES
    # ============================================================
    with menu_tabs[2]:  # Análises
        st.markdown('<div class="card-container">', unsafe_allow_html=True)
        st.markdown('<div class="card-title">📈 Análises Detalhadas</div>', unsafe_allow_html=True)

        col1, col2 = st.columns(2)

        with col1:
            df_mensal = pd.DataFrame({
                'Mês': ['Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez', 'Jan'],
                'Investimento': dados_mensais.values
            })
            fig_sazonal = px.area(
                df_mensal,
                x='Mês',
                y='Investimento',
                markers=True,
                title='Evolução do Investimento Mensal',
                color_discrete_sequence=[SPRES_BLUE]
            )
            fig_sazonal.update_traces(
                line=dict(color=SPRES_BLUE, width=3),
                marker=dict(color=SPRES_YELLOW, size=10, line=dict(color=SPRES_BLUE, width=2)),
                fillcolor='rgba(0, 75, 141, 0.12)'
            )
            fig_sazonal.update_layout(
                plot_bgcolor='rgba(0,0,0,0)',
                paper_bgcolor='rgba(0,0,0,0)',
                font=dict(color=SPRES_TEXT, family='Inter, sans-serif'),
                title_font_color=SPRES_BLUE_DARK,
                title_font_size=15,
                margin=dict(l=20, r=20, t=40, b=20),
                xaxis=dict(gridcolor='rgba(0,75,141,0.08)', showgrid=False),
                yaxis=dict(gridcolor='rgba(0,75,141,0.08)', tickprefix='R$ '),
                showlegend=False
            )
            st.plotly_chart(fig_sazonal, use_container_width=True, config={'displayModeBar': False})

        with col2:
            top_veiculos = distribuicao_veiculo.tail(5)
            fig_top = px.bar(
                x=top_veiculos.values,
                y=top_veiculos.index,
                orientation='h',
                title='Maiores Investimentos por Veículo',
                color=top_veiculos.values,
                color_continuous_scale=[SPRES_BLUE_LIGHT, SPRES_BLUE, SPRES_BLUE_DARK]
            )
            fig_top.update_traces(
                texttemplate='R$ %{x:,.0f}',
                textposition='inside',
                insidetextfont=dict(color=SPRES_WHITE)
            )
            fig_top.update_layout(
                xaxis_title='Investimento (R$)',
                yaxis_title=None,
                showlegend=False,
                plot_bgcolor='rgba(0,0,0,0)',
                paper_bgcolor='rgba(0,0,0,0)',
                font=dict(color=SPRES_TEXT, family='Inter, sans-serif'),
                title_font_color=SPRES_BLUE_DARK,
                title_font_size=15,
                margin=dict(l=20, r=30, t=40, b=20),
                coloraxis_showscale=False
            )
            st.plotly_chart(fig_top, use_container_width=True, config={'displayModeBar': False})

        st.markdown("---")
        st.markdown('<div class="card-title">📊 Análise de Controle</div>', unsafe_allow_html=True)

        col3, col4, col5 = st.columns(3)
        with col3:
            st.metric("✅ Total Controle Aprovado", formatar_moeda(big_numbers['total_controle_aprovado']))
        with col4:
            st.metric("🍊 Total Controle Utilizado", formatar_moeda(big_numbers['total_controle_utilizado']))
        with col5:
            st.metric("💰 Total Saldo Positivo", formatar_moeda(big_numbers['total_saldo_positivo']))

        st.markdown("---")
        st.markdown('<div class="card-title">📊 Aprovado vs Utilizado por Mês</div>', unsafe_allow_html=True)

        df_comp = pd.DataFrame({
            'Mês': ['Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez', 'Jan'],
            'Aprovado': [v for v in totais_aprovado_mensal.values()],
            'Utilizado': [v for v in totais_utilizado_mensal.values()],
            'Saldo': [v for v in totais_saldo_mensal.values()]
        })

        fig_comp = go.Figure()
        fig_comp.add_trace(go.Bar(
            name='Aprovado', x=df_comp['Mês'], y=df_comp['Aprovado'],
            marker_color=SPRES_BLUE, opacity=0.8
        ))
        fig_comp.add_trace(go.Bar(
            name='Utilizado', x=df_comp['Mês'], y=df_comp['Utilizado'],
            marker_color=SPRES_YELLOW, opacity=0.8
        ))
        fig_comp.add_trace(go.Scatter(
            name='Saldo', x=df_comp['Mês'], y=df_comp['Saldo'],
            mode='lines+markers',
            line=dict(color=SPRES_ORANGE, width=3),
            marker=dict(size=8, color=SPRES_ORANGE)
        ))
        fig_comp.update_layout(
            barmode='group',
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font=dict(color=SPRES_TEXT, family='Inter, sans-serif'),
            margin=dict(l=20, r=20, t=20, b=20),
            height=380,
            xaxis=dict(showgrid=False, gridcolor='rgba(0,75,141,0.08)'),
            yaxis=dict(gridcolor='rgba(0,75,141,0.08)', tickprefix='R$ ')
        )
        st.plotly_chart(fig_comp, use_container_width=True, config={'displayModeBar': False})

        st.markdown('</div>', unsafe_allow_html=True)

    # ============================================================
    # RODAPÉ
    # ============================================================
    data_source = "SharePoint Online" if (SHAREPOINT_AVAILABLE and GraphConfig.is_configured()) else "Dados Estáticos"
    st.markdown(f"""
    <div class="footer">
        <strong>Sucos Spres</strong> • Dados: {data_source} • Ciclo fev/26 – jan/27 •
        Gestão Mídia Pro
    </div>
    """, unsafe_allow_html=True)

# ============================================
# 7. TELA DE LOGIN — UNIFICADA (ESTRUTURA + EFEITOS)
# ============================================
def pagina_login():
    """Tela de login Spres — estrutura correta + efeitos animados de background."""
    
    # ===== CONTROLE: Se já estiver autenticado, redireciona =====
    if st.session_state.get("autenticado", False):
        st.rerun()
        return
    
    logo_b64 = logo_spres_base64()
    
    # ===== CSS GLOBAL =====
    st.markdown("""
    <style>
        /* ESCONDER elementos do Streamlit */
        #MainMenu {visibility: hidden;}
        header {visibility: hidden;}
        footer {visibility: hidden;}
        .stDeployButton {display: none;}
        [data-testid="stToolbar"] {display: none;}
        [data-testid="stDecoration"] {display: none;}
        [data-testid="stHeader"] {display: none !important;}
        .stAlert {display: none !important;}
        
        /* ===== BACKGROUND ANIMADO ===== */
        .stApp {
            background: linear-gradient(-45deg, #0C1B2E, #0A1F35, #00325F, #004B8D, #0A1F35);
            background-size: 400% 400%;
            animation: gradient-flow 16s ease infinite;
        }
 
        @keyframes gradient-flow {
            0% { background-position: 0% 50%; }
            50% { background-position: 100% 50%; }
            100% { background-position: 0% 50%; }
        }
        
        /* Centralizar verticalmente */
        .stAppViewContainer > section {
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
            min-height: 100vh !important;
        }
        
        /* Container principal */
        .block-container {
            max-width: 520px !important;
            padding: 0 20px !important;
            margin: 0 auto !important;
        }
        
        /* ===== PARTÍCULAS FLUTUANTES ===== */
        .login-particles {
            position: fixed;
            top: 0;
            left: 0;
            width: 100vw;
            height: 100vh;
            overflow: hidden;
            pointer-events: none;
            z-index: 0;
        }
        .login-particle {
            position: absolute;
            bottom: -60px;
            border-radius: 50%;
            opacity: 0;
            animation: float-up linear infinite;
        }
        @keyframes float-up {
            0% { transform: translateY(0) translateX(0) scale(0.6); opacity: 0; }
            10% { opacity: 0.5; }
            90% { opacity: 0.35; }
            100% { transform: translateY(-110vh) translateX(30px) scale(1); opacity: 0; }
        }
        
        /* ===== CARD COM BORDA ANIMADA (SÓ A BORDA MUDA DE COR) ===== */
        .login-card-wrapper {
            position: relative;
            border-radius: 24px;
            padding: 3px;
            background: linear-gradient(90deg, #FFD600, #2E7DD1, #FFD600);
            background-size: 200% 100%;
            animation: border-glow 3s ease infinite;
            box-shadow: 0 8px 40px rgba(0, 0, 0, 0.35), 0 0 40px rgba(46, 125, 209, 0.15);
        }
        
        @keyframes border-glow {
            0% { background-position: 0% 50%; }
            50% { background-position: 100% 50%; }
            100% { background-position: 0% 50%; }
        }
        
        .login-card-inner {
            background: rgba(12, 27, 46, 0.95);
            backdrop-filter: blur(24px);
            -webkit-backdrop-filter: blur(24px);
            border-radius: 22px;
            padding: 44px 48px 36px 48px;
            position: relative;
            overflow: hidden;
            animation: card-enter 0.7s cubic-bezier(0.16, 1, 0.3, 1);
        }
        
        /* Barra shimmer no topo do card (azul/amarelo) */
        .login-card-inner::before {
            content: '';
            position: absolute;
            top: 0;
            left: -100%;
            right: 0;
            height: 3px;
            width: 300%;
            background: linear-gradient(90deg, transparent, #2E7DD1, #FFD600, #2E7DD1, transparent);
            animation: shimmer-bar 4s linear infinite;
        }
        
        @keyframes shimmer-bar {
            0% { transform: translateX(0); }
            100% { transform: translateX(33.33%); }
        }
 
        @keyframes card-enter {
            0% { opacity: 0; transform: translateY(24px) scale(0.97); }
            100% { opacity: 1; transform: translateY(0) scale(1); }
        }
        
        /* ===== LOGO (COR FIXA, NÃO MUDA) ===== */
        .login-logo {
            text-align: center;
            margin-bottom: 10px;
            animation: logo-enter 0.8s cubic-bezier(0.16, 1, 0.3, 1) 0.15s both;
        }
 
        @keyframes logo-enter {
            0% { opacity: 0; transform: translateY(-10px) scale(0.9); }
            100% { opacity: 1; transform: translateY(0) scale(1); }
        }
        
        .login-logo img {
            max-width: 160px;
            height: auto;
            filter: drop-shadow(0 6px 18px rgba(0,0,0,0.35));
        }
        
        /* Subtítulo animado */
        .login-subtitle {
            text-align: center;
            color: rgba(255, 255, 255, 0.45);
            font-size: 11px;
            font-weight: 600;
            letter-spacing: 2.5px;
            text-transform: uppercase;
            margin-bottom: 6px;
            animation: fade-in-up 0.6s ease 0.3s both;
        }
        
        /* Título animado */
        .login-title {
            text-align: center;
            color: #FFFFFF;
            font-size: 26px;
            font-weight: 700;
            margin-bottom: 28px;
            letter-spacing: -0.3px;
            animation: fade-in-up 0.6s ease 0.4s both;
        }
 
        @keyframes fade-in-up {
            0% { opacity: 0; transform: translateY(8px); }
            100% { opacity: 1; transform: translateY(0); }
        }
        
        .login-title span {
            color: #FFD600;
            text-shadow: 0 0 24px rgba(255, 214, 0, 0.35);
        }
        
        /* Formulário animado */
        div[data-testid="stForm"] {
            animation: fade-in-up 0.6s ease 0.5s both;
            border: none !important;
            padding: 0 !important;
            margin: 0 !important;
            background: transparent !important;
        }
        
        div[data-testid="stForm"] > div {
            gap: 0px !important;
        }
        
        div[data-testid="stFormSubmitButton"] {
            margin-top: 6px !important;
        }
        
        /* ===== INPUTS — FUNDO ESCURO FIXO ===== */
        .stTextInput {
            margin-bottom: 16px !important;
        }
        
        .stTextInput > div {
            width: 100% !important;
        }
        
        .stTextInput label {
            color: rgba(255, 255, 255, 0.55) !important;
            font-size: 13px !important;
            font-weight: 500 !important;
            margin-bottom: 6px !important;
            padding-left: 2px !important;
            letter-spacing: 0.3px !important;
        }
        
        .stTextInput input {
            background: rgba(10, 25, 45, 0.85) !important;
            background-color: rgba(10, 25, 45, 0.85) !important;
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
            border: 1px solid rgba(255, 255, 255, 0.15) !important;
            border-radius: 12px !important;
            font-size: 15px !important;
            font-family: 'Inter', sans-serif !important;
            padding: 13px 16px !important;
            height: 50px !important;
            transition: all 0.25s ease !important;
            caret-color: #FFD600 !important;
            box-shadow: inset 0 1px 2px rgba(0,0,0,0.2) !important;
        }
        
        .stTextInput input::placeholder {
            color: rgba(255, 255, 255, 0.3) !important;
            font-size: 14px !important;
        }
        
        .stTextInput input:hover {
            background: rgba(10, 25, 45, 0.9) !important;
            border-color: rgba(255, 255, 255, 0.3) !important;
        }
        
        .stTextInput input:focus {
            border-color: #2E7DD1 !important;
            box-shadow: 0 0 0 3px rgba(46, 125, 209, 0.15), inset 0 1px 2px rgba(0,0,0,0.2) !important;
            background: rgba(10, 25, 45, 0.95) !important;
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
        }
        
        /* ===== BOTÃO ===== */
        .stButton {
            margin-top: 6px !important;
        }
        
        .stButton button {
            background: linear-gradient(135deg, #004B8D 0%, #2E7DD1 100%) !important;
            color: #FFFFFF !important;
            border: none !important;
            border-radius: 12px !important;
            font-size: 15px !important;
            font-weight: 600 !important;
            font-family: 'Inter', sans-serif !important;
            padding: 14px 24px !important;
            height: 52px !important;
            width: 100% !important;
            transition: all 0.25s cubic-bezier(0.4,0,0.2,1) !important;
            letter-spacing: 0.2px !important;
            position: relative !important;
            overflow: hidden !important;
        }
        
        .stButton button::before {
            content: '';
            position: absolute;
            top: 0;
            left: -75%;
            width: 50%;
            height: 100%;
            background: linear-gradient(120deg, transparent, rgba(255,255,255,0.35), transparent);
            transform: skewX(-20deg);
            transition: left 0.6s ease;
        }
        
        .stButton button:hover::before {
            left: 125%;
        }
        
        .stButton button:hover {
            transform: translateY(-2px) !important;
            box-shadow: 0 8px 28px rgba(0, 75, 141, 0.45), 0 0 0 1px rgba(255,214,0,0.25) !important;
        }
        
        .stButton button:active {
            transform: translateY(0) !important;
        }
        
        /* ===== MENSAGEM DE ERRO ===== */
        .login-error-msg {
            background: rgba(255, 59, 48, 0.12);
            border: 1px solid rgba(255, 59, 48, 0.2);
            border-radius: 12px;
            color: #FF7B7B;
            padding: 12px 16px;
            text-align: center;
            margin-top: 14px;
            font-size: 13px;
            font-weight: 500;
            animation: shake 0.4s ease;
        }
        
        @keyframes shake {
            0%, 100% { transform: translateX(0); }
            20% { transform: translateX(-5px); }
            40% { transform: translateX(5px); }
            60% { transform: translateX(-3px); }
            80% { transform: translateX(3px); }
        }
        
        /* ===== FOOTER ===== */
        .login-footer {
            text-align: center;
            color: rgba(255, 255, 255, 0.2);
            font-size: 11px;
            margin-top: 22px;
            letter-spacing: 0.3px;
            animation: fade-in-up 0.6s ease 0.6s both;
        }
        
        .login-footer strong {
            color: rgba(255, 214, 0, 0.4);
        }
        
        /* Tentativas */
        .login-attempts {
            text-align: center;
            color: rgba(255, 255, 255, 0.3);
            font-size: 11px;
            margin-top: 10px;
        }
    </style>
    """, unsafe_allow_html=True)
    
    # ===== PARTÍCULAS FLUTUANTES =====
    particulas_cfg = [
        (6, "5%", 46, "rgba(46,125,209,0.5)"),
        (10, "15%", 38, "rgba(255,214,0,0.45)"),
        (14, "27%", 62, "rgba(255,255,255,0.35)"),
        (8, "42%", 30, "rgba(255,138,30,0.4)"),
        (18, "55%", 50, "rgba(46,125,209,0.4)"),
        (11, "68%", 34, "rgba(255,214,0,0.4)"),
        (16, "80%", 44, "rgba(255,255,255,0.3)"),
        (9, "90%", 28, "rgba(255,138,30,0.35)"),
    ]
    particulas_html = ""
    for i, (dur, left, size, color) in enumerate(particulas_cfg):
        delay = i * 1.7
        particulas_html += (
            f'<div class="login-particle" style="left:{left}; width:{size}px; height:{size}px; '
            f'background:{color}; animation-duration:{dur}s; animation-delay:{delay}s;"></div>'
        )
    
    st.markdown(f"""
    <div class="login-particles">{particulas_html}</div>
    """, unsafe_allow_html=True)
    
    # ===== ESTADO =====
    if "login_error" not in st.session_state:
        st.session_state.login_error = ""
    if "login_attempts" not in st.session_state:
        st.session_state.login_attempts = 0
    
    # ===== LAYOUT: Colunas para centralizar =====
    col_esq, col_center, col_dir = st.columns([1, 2.8, 1])
    
    with col_center:
        # ===== CARD COM BORDA GRADIENTE (HTML PURO) =====
        logo_html = f'<img src="data:image/png;base64,{logo_b64}" alt="Sucos Spres">' if logo_b64 else "🍊"
        
        st.markdown(f"""
        <div class="login-card-wrapper">
            <div class="login-card-inner">
                <div class="login-logo">
                    {logo_html}
                </div>
                <div class="login-subtitle">Portal de Acesso</div>
                <div class="login-title">
                    Gestão <span>Mídia</span> Pro
                </div>
        """, unsafe_allow_html=True)
        
        # ===== FORMULÁRIO (Streamlit dentro do HTML) =====
        with st.form(key="login_form", clear_on_submit=False, border=False):
            usuario = st.text_input("Usuário", placeholder="👤  Digite seu usuário", key="login_user")
            senha = st.text_input("Senha", type="password", placeholder="🔒  Digite sua senha", key="login_pass")
            
            submitted = st.form_submit_button("🚀  Acessar Dashboard", use_container_width=True)
            
            if submitted:
                if not usuario.strip() or not senha.strip():
                    st.session_state.login_error = "⚠️ Preencha usuário e senha."
                    st.session_state.login_attempts += 1
                    st.rerun()
                elif usuario in USUARIOS and USUARIOS[usuario] == senha:
                    # ===== LOGIN RÁPIDO =====
                    st.session_state["autenticado"] = True
                    st.session_state["usuario"] = usuario
                    st.session_state["login_time"] = datetime.now()
                    st.session_state.login_error = ""
                    st.session_state.login_attempts = 0
                    st.session_state.dados_carregados = False
                    st.rerun()
                    return
                else:
                    st.session_state.login_error = "❌ Usuário ou senha inválidos."
                    st.session_state.login_attempts += 1
                    st.rerun()
        
        # ===== FECHAR CARD E ADICIONAR ERRO/FOOTER =====
        error_html = ""
        if st.session_state.login_error:
            error_html = f'<div class="login-error-msg">{st.session_state.login_error}</div>'
        
        attempts_html = ""
        if st.session_state.login_attempts >= 3:
            attempts_html = f'<div class="login-attempts">Tentativa {st.session_state.login_attempts}/5</div>'
        
        st.markdown(f"""
                {error_html}
                {attempts_html}
                <div class="login-footer">
                    <strong>Sucos Spres</strong> • Desde 1991
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

# ============================================
# ENTRADA DO APLICATIVO
# ============================================
if __name__ == "__main__":
    # Inicializa session_state
    if "autenticado" not in st.session_state:
        st.session_state.autenticado = False
    if "dados_carregados" not in st.session_state:
        st.session_state.dados_carregados = False
    
    # Verifica autenticação
    if st.session_state.autenticado:
        # Se estiver autenticado, mostra o dashboard
        main()
    else:
        # Se não estiver, mostra a tela de login
        pagina_login()