# ============================================
# excel_integration.py - Módulo para integração com Excel
# Suporte: SharePoint Online + Arquivo Local + Upload
# ============================================
import pandas as pd
import streamlit as st
from pathlib import Path
import io
import time
import logging
from datetime import datetime
import os
import re

# ============================================
# CARREGAR VARIÁVEIS DE AMBIENTE
# ============================================
from dotenv import load_dotenv

env_paths = [
    ".env",
    os.path.join(os.path.dirname(__file__), ".env"),
    "../.env",
]

env_loaded = False
for path in env_paths:
    if os.path.exists(path):
        load_dotenv(path)
        env_loaded = True
        print(f"✅ Arquivo .env carregado de: {path}")
        break

if not env_loaded:
    print("⚠️ Arquivo .env não encontrado. Usando variáveis de ambiente do sistema.")
    load_dotenv()

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
    
    GraphConfig.TENANT_ID = os.getenv("AZURE_TENANT_ID", "")
    GraphConfig.CLIENT_ID = os.getenv("AZURE_CLIENT_ID", "")
    GraphConfig.CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET", "")
    GraphConfig.SHAREPOINT_SITE_URL = os.getenv("SHAREPOINT_SITE_URL", "")
    GraphConfig.EXCEL_FILENAME = os.getenv("EXCEL_FILENAME", "base_spres_projeto_refatorada.xlsx")
    GraphConfig.CACHE_MINUTES = int(os.getenv("CACHE_MINUTES", "30"))
    
    print(f"📋 Configurações do SharePoint carregadas:")
    print(f"  Site URL: {GraphConfig.SHAREPOINT_SITE_URL[:50]}..." if GraphConfig.SHAREPOINT_SITE_URL else "  Site URL: ❌ Não configurado")
    print(f"  Arquivo: {GraphConfig.EXCEL_FILENAME}")
    print(f"  Cache: {GraphConfig.CACHE_MINUTES} min")
    
except ImportError as e:
    SHAREPOINT_AVAILABLE = False
    print(f"⚠️ Módulo graph_api não disponível: {str(e)}")

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class ExcelIntegration:
    
    def __init__(self):
        self.dados_geral = None
        self.dados_controle = None
        self.dados_extra = None
        self.big_numbers = None
        self.totais_aprovado_mensal = None
        self.totais_utilizado_mensal = None
        self.totais_saldo_mensal = None
        self.dados_estudo = None
        self.origem_dados = "Não carregado"
        self.ultima_atualizacao = None
        self.meses = ['fev/26', 'mar/26', 'abr/26', 'mai/26', 'jun/26',
                     'jul/26', 'ago/26', 'set/26', 'out/26', 'nov/26', 'dez/26', 'jan/27']
        
        self._sharepoint_configurado = SHAREPOINT_AVAILABLE and GraphConfig.is_configured()
        if self._sharepoint_configurado:
            logger.info("SharePoint configurado e disponível")
        else:
            logger.warning("SharePoint não configurado ou indisponível")
    
    def _converter_valor(self, valor):
        """
        Converte valor monetário para float de forma segura.
        CORRIGIDO: Suporta formato brasileiro (R$ 1.234,56)
        """
        if valor is None or valor == '' or pd.isna(valor):
            return 0.0
        
        if isinstance(valor, (int, float)):
            return float(valor)
        
        if isinstance(valor, str):
            # Remove R$, espaços e formatação
            valor_limpo = valor.replace('R$', '').replace(' ', '').replace('.', '').replace(',', '.')
            
            # Remove qualquer caractere que não seja número ou ponto
            import re
            valor_limpo = re.sub(r'[^\d.]', '', valor_limpo)
            
            if not valor_limpo:
                return 0.0
            
            try:
                return float(valor_limpo)
            except ValueError:
                return 0.0
        
        return 0.0
    
    def carregar_sharepoint(self, forcar: bool = False) -> bool:
        if not SHAREPOINT_AVAILABLE:
            logger.warning("Módulo SharePoint não disponível")
            return False
        
        if not GraphConfig.is_configured():
            logger.warning("Configuração do SharePoint incompleta")
            return False
        
        try:
            if forcar:
                logger.info("🔄 Forçando recarga (ignorando cache)...")
                try:
                    invalidate_sharepoint_cache()
                    st.cache_data.clear()
                except:
                    pass
                time.sleep(0.5)
            
            logger.info("Iniciando carga do SharePoint...")
            data = load_sharepoint_data()
            
            if "error" in data:
                logger.error(f"Erro no SharePoint: {data['error']}")
                return False
            
            self._processar_dados_sharepoint(data)
            self.origem_dados = "SharePoint Online" + (" (Forçado)" if forcar else "")
            self.ultima_atualizacao = datetime.now()
            
            logger.info("✅ Dados carregados com sucesso do SharePoint")
            return True
                
        except Exception as e:
            logger.error(f"Erro ao carregar SharePoint: {str(e)}")
            return False
    
    def _processar_dados_sharepoint(self, data: dict):
        """Processa os dados do SharePoint para o formato do dashboard"""
        
        df_controle = data.get("controle", pd.DataFrame())
        df_extra = data.get("extra", pd.DataFrame())
        df_estudo = data.get("estudo", pd.DataFrame())
        
        meses_disponiveis = []
        for mes in self.meses:
            if mes in df_controle.columns:
                meses_disponiveis.append(mes)
        
        if not meses_disponiveis:
            logger.warning("Nenhum mês encontrado no DataFrame. Usando meses padrão.")
            meses_disponiveis = self.meses
        
        logger.info(f"Meses disponíveis: {len(meses_disponiveis)} meses")
        
        dados_geral = {}
        total_geral = 0
        
        for mes in self.meses:
            if mes in df_controle.columns:
                try:
                    valores = df_controle[mes].apply(self._converter_valor)
                    valor = valores.sum()
                except Exception as e:
                    logger.warning(f"Erro ao converter mês {mes}: {str(e)}")
                    valor = 0.0
                
                dados_geral[mes] = float(valor) if not pd.isna(valor) else 0
                total_geral += dados_geral[mes]
            else:
                dados_geral[mes] = 0
        
        dados_geral['total'] = total_geral
        self.dados_geral = dados_geral
        
        for mes in meses_disponiveis:
            if mes in df_controle.columns:
                df_controle[mes] = df_controle[mes].apply(self._converter_valor)
        
        if 'total' not in df_controle.columns and meses_disponiveis:
            df_controle['total'] = df_controle[meses_disponiveis].sum(axis=1)
        
        self.dados_controle = df_controle
        
        try:
            self.big_numbers = get_sharepoint_big_numbers()
        except Exception as e:
            logger.warning(f"Erro ao obter big numbers: {str(e)}")
            self.big_numbers = self._extrair_big_numbers(df_controle)
        
        try:
            totais_mensais = get_sharepoint_totais_mensais()
            self.totais_aprovado_mensal = totais_mensais.get('aprovado', {})
            self.totais_utilizado_mensal = totais_mensais.get('utilizado', {})
            self.totais_saldo_mensal = totais_mensais.get('saldo', {})
        except Exception as e:
            logger.warning(f"Erro ao obter totais mensais: {str(e)}")
            self.totais_aprovado_mensal = {mes: 0 for mes in meses_disponiveis}
            self.totais_utilizado_mensal = {mes: 0 for mes in meses_disponiveis}
            self.totais_saldo_mensal = {mes: 0 for mes in meses_disponiveis}
        
        self.dados_estudo = {
            'aba_geral': {},
            'aba_geral_controle': {},
            'aba_geral_extra': {}
        }
        
        if not df_estudo.empty:
            for idx, row in df_estudo.iterrows():
                mes_ref = str(row.iloc[0]).strip()
                if mes_ref and '/' in mes_ref:
                    mes_key = mes_ref.lower()
                    if mes_key in self.meses:
                        self.dados_estudo['aba_geral'][mes_key] = self._converter_valor(row.iloc[1]) if len(row) > 1 else 0
                        self.dados_estudo['aba_geral_controle'][mes_key] = self._converter_valor(row.iloc[2]) if len(row) > 2 else 0
                        self.dados_estudo['aba_geral_extra'][mes_key] = self._converter_valor(row.iloc[3]) if len(row) > 3 else 0
        
        if df_extra.empty:
            self.dados_extra = pd.DataFrame()
        else:
            for mes in meses_disponiveis:
                if mes in df_extra.columns:
                    df_extra[mes] = df_extra[mes].apply(self._converter_valor)
            self.dados_extra = df_extra
        
        logger.info(f"Dados processados: {len(df_controle)} linhas")
    
    def _extrair_big_numbers(self, df_controle):
        big_numbers = {
            'total_controle_aprovado': 0,
            'total_controle_utilizado': 0,
            'total_saldo_positivo': 0
        }
        
        for idx, row in df_controle.iterrows():
            desc = str(row.iloc[0]).upper() if row.iloc[0] else ''
            
            # ===== BUSCA POR PALAVRAS-CHAVE =====
            if 'APROV' in desc or 'APROVADO' in desc:
                big_numbers['total_controle_aprovado'] = self._converter_valor(row.iloc[-1])
            elif 'UTILIZ' in desc or 'UTILIZADO' in desc:
                big_numbers['total_controle_utilizado'] = self._converter_valor(row.iloc[-1])
            elif 'SALDO' in desc or 'POSITIVO' in desc:
                big_numbers['total_saldo_positivo'] = self._converter_valor(row.iloc[-1])
        
        return big_numbers
    
    def carregar_excel_local(self, caminho_arquivo: str) -> bool:
        try:
            if not Path(caminho_arquivo).exists():
                logger.error(f"Arquivo não encontrado: {caminho_arquivo}")
                return False
            
            logger.info(f"Carregando arquivo local: {caminho_arquivo}")
            xls = pd.ExcelFile(caminho_arquivo)
            
            if 'GERAL' in xls.sheet_names:
                df_geral = pd.read_excel(caminho_arquivo, sheet_name='GERAL')
                self.dados_geral = self._processar_aba_geral(df_geral)
            
            if 'GERAL CONTROLE (2) NOVA' in xls.sheet_names:
                df_controle = pd.read_excel(caminho_arquivo, sheet_name='GERAL CONTROLE (2) NOVA')
                self.dados_controle = self._processar_aba_controle(df_controle)
            
            if 'GERAL EXTRA' in xls.sheet_names:
                self.dados_extra = pd.read_excel(caminho_arquivo, sheet_name='GERAL EXTRA')
            
            if 'ESTUDO' in xls.sheet_names:
                df_estudo = pd.read_excel(caminho_arquivo, sheet_name='ESTUDO')
                self.dados_estudo = self._processar_aba_estudo(df_estudo)
            
            self.origem_dados = f"Arquivo Local: {Path(caminho_arquivo).name}"
            self.ultima_atualizacao = datetime.now()
            
            logger.info(f"✅ Dados carregados do arquivo local: {caminho_arquivo}")
            return True
            
        except Exception as e:
            logger.error(f"Erro ao carregar arquivo local: {str(e)}")
            return False
    
    def _processar_aba_geral(self, df):
        dados = {}
        for mes in self.meses:
            if mes in df.columns:
                dados[mes] = self._converter_valor(df[mes].iloc[-1])
        dados['total'] = sum(dados.values())
        return dados
    
    def _processar_aba_controle(self, df):
        df = df.copy()
        if len(df.columns) >= 16:
            df.columns = ['veiculo', 'canal', 'obs'] + self.meses + ['total']
        return df
    
    def _processar_aba_estudo(self, df):
        dados_estudo = {'aba_geral': {}, 'aba_geral_controle': {}, 'aba_geral_extra': {}}
        
        for idx, row in df.iterrows():
            mes_ref = str(row.iloc[0]).strip()
            if mes_ref and '/' in mes_ref:
                mes_key = mes_ref.lower()
                if mes_key in self.meses:
                    dados_estudo['aba_geral'][mes_key] = self._converter_valor(row.iloc[1]) if len(row) > 1 else 0
                    dados_estudo['aba_geral_controle'][mes_key] = self._converter_valor(row.iloc[2]) if len(row) > 2 else 0
                    dados_estudo['aba_geral_extra'][mes_key] = self._converter_valor(row.iloc[3]) if len(row) > 3 else 0
        
        return dados_estudo
    
    def carregar_excel_upload(self, arquivo_upload) -> bool:
        try:
            if arquivo_upload is None:
                return False
            
            logger.info(f"Carregando arquivo via upload: {arquivo_upload.name}")
            
            import tempfile
            with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
                tmp.write(arquivo_upload.getvalue())
                tmp_path = tmp.name
            
            resultado = self.carregar_excel_local(tmp_path)
            
            try:
                Path(tmp_path).unlink()
            except:
                pass
            
            if resultado:
                self.origem_dados = f"Upload: {arquivo_upload.name}"
                self.ultima_atualizacao = datetime.now()
            
            return resultado
            
        except Exception as e:
            logger.error(f"Erro ao processar upload: {str(e)}")
            return False
    
    def get_dados_atualizados(self):
        if self.dados_geral is None or self.dados_controle is None:
            if SHAREPOINT_AVAILABLE and GraphConfig.is_configured():
                if self.carregar_sharepoint():
                    return self._montar_retorno()
            
            logger.warning("Usando dados estáticos como fallback")
            return self._carregar_dados_estaticos()
        
        return self._montar_retorno()
    
    def _montar_retorno(self):
        """Monta a tupla de retorno com todos os dados"""
        # Verifica se dados_extra é None ou vazio de forma segura
        if self.dados_extra is None or (isinstance(self.dados_extra, pd.DataFrame) and self.dados_extra.empty):
            dados_extra = pd.DataFrame()
        else:
            dados_extra = self.dados_extra
        
        # Verifica se dados_estudo é None ou vazio
        if self.dados_estudo is None or not self.dados_estudo:
            dados_estudo = {'aba_geral': {}, 'aba_geral_controle': {}, 'aba_geral_extra': {}}
        else:
            dados_estudo = self.dados_estudo
        
        return (
            self.dados_geral,
            self.dados_controle,
            self.big_numbers or {
                'total_controle_aprovado': 0,
                'total_controle_utilizado': 0,
                'total_saldo_positivo': 0
            },
            self.totais_aprovado_mensal or {mes: 0 for mes in self.meses},
            self.totais_utilizado_mensal or {mes: 0 for mes in self.meses},
            self.totais_saldo_mensal or {mes: 0 for mes in self.meses},
            dados_estudo,
            dados_extra
        )
    
    def _carregar_dados_estaticos(self):
        try:
            from app import carregar_dados_estaticos
            return carregar_dados_estaticos()
        except:
            meses = ['fev/26', 'mar/26', 'abr/26', 'mai/26', 'jun/26',
                     'jul/26', 'ago/26', 'set/26', 'out/26', 'nov/26', 'dez/26', 'jan/27']
            dados_geral = {mes: 0 for mes in meses}
            dados_geral['total'] = 0
            dados_controle = pd.DataFrame()
            big_numbers = {'total_controle_aprovado': 0, 'total_controle_utilizado': 0, 'total_saldo_positivo': 0}
            totais = {mes: 0 for mes in meses}
            dados_estudo = {'aba_geral': {}, 'aba_geral_controle': {}, 'aba_geral_extra': {}}
            dados_extra = pd.DataFrame()
            return (dados_geral, dados_controle, big_numbers, totais, totais, totais, dados_estudo, dados_extra)
    
    def get_status(self) -> dict:
        return {
            'origem': self.origem_dados,
            'ultima_atualizacao': self.ultima_atualizacao.strftime('%d/%m/%Y %H:%M:%S') if self.ultima_atualizacao else 'Nunca',
            'sharepoint_disponivel': SHAREPOINT_AVAILABLE,
            'sharepoint_configurado': GraphConfig.is_configured() if SHAREPOINT_AVAILABLE else False,
            'dados_geral': len(self.dados_geral) if self.dados_geral else 0,
            'dados_controle': len(self.dados_controle) if self.dados_controle is not None else 0,
        }
    
    def forcar_atualizacao(self):
        self.ultima_atualizacao = None
        logger.info("Forçando atualização de dados...")
        
        if SHAREPOINT_AVAILABLE and GraphConfig.is_configured():
            try:
                invalidate_sharepoint_cache()
                st.cache_data.clear()
                logger.info("Cache invalidado")
            except Exception as e:
                logger.error(f"Erro ao invalidar cache: {str(e)}")

# ============================================
# FUNÇÕES DE UTILIDADE PARA O DASHBOARD
# ============================================
def criar_interface_upload():
    with st.expander("📤 Carregar Planilha Manualmente"):
        uploaded_file = st.file_uploader(
            "Selecione uma planilha Excel",
            type=['xlsx', 'xls'],
            help="Faça upload de uma planilha para substituir os dados atuais"
        )
        
        if uploaded_file is not None:
            if st.button("📊 Processar Planilha", use_container_width=True):
                integracao = ExcelIntegration()
                if integracao.carregar_excel_upload(uploaded_file):
                    st.success(f"✅ Planilha carregada com sucesso! Origem: {integracao.origem_dados}")
                    st.rerun()
                else:
                    st.error("❌ Erro ao processar a planilha. Verifique o formato.")

def mostrar_status_integracao():
    integracao = ExcelIntegration()
    status = integracao.get_status()
    
    st.markdown("### 📊 Status da Integração")
    
    if status['sharepoint_disponivel'] and status['sharepoint_configurado']:
        try:
            success, message = test_sharepoint_connection()
            if success:
                st.success("✅ SharePoint Online Conectado")
                st.caption(f"📁 {GraphConfig.EXCEL_FILENAME}")
                st.caption(f"🔄 Cache: {GraphConfig.CACHE_MINUTES} min")
            else:
                st.warning(f"⚠️ {message}")
        except Exception as e:
            st.warning(f"⚠️ SharePoint: Erro na conexão - {str(e)}")
    else:
        st.info("📌 SharePoint não configurado")
        st.caption("Configure o .env com as credenciais")
    
    st.caption(f"🔄 Última atualização: {status['ultima_atualizacao']}")
    st.caption(f"📁 Origem: {status['origem']}")
    
    if status['dados_controle'] > 0:
        st.caption(f"📊 Itens carregados: {status['dados_controle']}")
    
    if st.button("🔄 Recarregar Dados", use_container_width=True):
        integracao.forcar_atualizacao()
        if integracao.carregar_sharepoint(forcar=True):
            st.success("✅ Dados recarregados com sucesso!")
        else:
            st.warning("⚠️ Falha ao recarregar. Usando dados em cache.")
        time.sleep(0.5)
        st.rerun()

# ============================================
# TESTE RÁPIDO
# ============================================
if __name__ == "__main__":
    print("\n=== TESTE DA INTEGRAÇÃO EXCEL ===\n")
    
    print("📋 Configurações carregadas:")
    print(f"  SharePoint disponível: {SHAREPOINT_AVAILABLE}")
    if SHAREPOINT_AVAILABLE:
        print(f"  SharePoint configurado: {GraphConfig.is_configured()}")
        print(f"  Site URL: {GraphConfig.SHAREPOINT_SITE_URL[:50]}..." if GraphConfig.SHAREPOINT_SITE_URL else "  Site URL: ❌")
        print(f"  Arquivo: {GraphConfig.EXCEL_FILENAME}")
        print(f"  Cache: {GraphConfig.CACHE_MINUTES} min")
    
    print("\n🚀 Testando integração...")
    integracao = ExcelIntegration()
    
    if SHAREPOINT_AVAILABLE and GraphConfig.is_configured():
        if integracao.carregar_sharepoint():
            print("✅ Dados carregados do SharePoint com sucesso!")
            status = integracao.get_status()
            print(f"  Origem: {status['origem']}")
            print(f"  Itens: {status['dados_controle']}")
            print(f"  Última atualização: {status['ultima_atualizacao']}")
        else:
            print("❌ Falha ao carregar dados do SharePoint")
    else:
        print("⚠️ SharePoint não disponível. Teste apenas com arquivo local.")
        arquivo = "base_spres_projeto_refatorada.xlsx"
        if Path(arquivo).exists():
            if integracao.carregar_excel_local(arquivo):
                print(f"✅ Dados carregados do arquivo local: {arquivo}")
                status = integracao.get_status()
                print(f"  Origem: {status['origem']}")
                print(f"  Itens: {status['dados_controle']}")
            else:
                print(f"❌ Arquivo não encontrado: {arquivo}")
        else:
            print(f"⚠️ Arquivo não encontrado: {arquivo}")