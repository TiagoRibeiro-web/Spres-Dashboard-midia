# ============================================
# graph_api.py - Conexão Microsoft Graph API
# ============================================
import os
import requests
import pandas as pd
import streamlit as st
from datetime import datetime, timedelta
from typing import Optional, Dict, Any, Tuple, List
import json
import hashlib
import time
from urllib.parse import urlparse
import logging
import re

# ============================================
# CONTROLE DE CARREGAMENTO LAZY
# ============================================
# Flag para controlar se o módulo já foi inicializado
_initialized = False

def initialize_graph():
    """Inicializa o módulo Graph API apenas quando necessário"""
    global _initialized
    if not _initialized:
        logger.info("🔄 Inicializando Graph API...")
        _initialized = True

# ============================================
# CARREGAR VARIÁVEIS DE AMBIENTE / SECRETS
# ============================================
from dotenv import load_dotenv

def get_secret(key, default=None):
    """Obtém uma variável do secrets.toml (Cloud) ou .env (local)"""
    # Tenta do secrets.toml (Streamlit Cloud)
    try:
        if hasattr(st, 'secrets') and key in st.secrets:
            return st.secrets[key]
    except:
        pass
    
    # Tenta do .env (desenvolvimento local)
    value = os.getenv(key)
    if value is not None:
        return value
    
    return default

# Tenta carregar o .env de diferentes locais (apenas para desenvolvimento local)
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

# Configuração de logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ============================================
# CONFIGURAÇÃO
# ============================================
class GraphConfig:
    """Configurações da Microsoft Graph API"""
    # AGORA USA get_secret() em vez de os.getenv()
    TENANT_ID = get_secret("AZURE_TENANT_ID", "")
    CLIENT_ID = get_secret("AZURE_CLIENT_ID", "")
    CLIENT_SECRET = get_secret("AZURE_CLIENT_SECRET", "")
    SHAREPOINT_SITE_URL = get_secret("SHAREPOINT_SITE_URL", "")
    EXCEL_FILENAME = get_secret("EXCEL_FILENAME", "base_spres_projeto_refatorada.xlsx")
    CACHE_MINUTES = int(get_secret("CACHE_MINUTES", "30"))
    
    # Mapeamento de abas e ranges do Excel
    WORKSHEET_MAPPING = {
        "geral": {"name": "GERAL", "range": "A1:M2"},
        "controle": {"name": "GERAL CONTROLE (2) NOVA", "range": "A1:O50"},
        "estudo": {"name": "ESTUDO", "range": "A1:M12"},
        "extra": {"name": "GERAL EXTRA", "range": "A1:N7"}
    }
    
    @classmethod
    def is_configured(cls) -> bool:
        """Verifica se todas as variáveis obrigatórias estão preenchidas"""
        return all([
            cls.TENANT_ID,
            cls.CLIENT_ID,
            cls.CLIENT_SECRET,
            cls.SHAREPOINT_SITE_URL
        ])
    
    @classmethod
    def get_cache_key(cls, worksheet_key: str) -> str:
        """Gera chave de cache para uma aba específica"""
        return f"graph_cache_{worksheet_key}_{cls.EXCEL_FILENAME}"
    
    @classmethod
    def is_onedrive_url(cls) -> bool:
        """Verifica se a URL é do OneDrive pessoal"""
        return "-my." in cls.SHAREPOINT_SITE_URL or "personal/" in cls.SHAREPOINT_SITE_URL
    
    @classmethod
    def should_cache(cls) -> bool:
        """Verifica se o cache está habilitado"""
        return cls.CACHE_MINUTES > 0

# ============================================
# CACHE EM MEMÓRIA
# ============================================
class GraphCache:
    """Sistema de cache em memória com expiração"""
    
    def __init__(self):
        self._cache: Dict[str, Tuple[pd.DataFrame, datetime]] = {}
        self._file_timestamp: Optional[datetime] = None
        self._last_clear: Optional[datetime] = None
    
    def get(self, key: str) -> Optional[pd.DataFrame]:
        """Recupera dados do cache se ainda válidos"""
        if not GraphConfig.should_cache():
            logger.debug(f"Cache desabilitado para {key}")
            return None
            
        if key in self._cache:
            data, timestamp = self._cache[key]
            if datetime.now() - timestamp < timedelta(minutes=GraphConfig.CACHE_MINUTES):
                logger.info(f"Cache hit para {key}")
                return data.copy()
            else:
                logger.info(f"Cache expirado para {key}")
                del self._cache[key]
        return None
    
    def set(self, key: str, data: pd.DataFrame):
        """Armazena dados no cache com timestamp"""
        if not GraphConfig.should_cache():
            logger.debug(f"Cache desabilitado, não armazenando {key}")
            return
            
        self._cache[key] = (data.copy(), datetime.now())
        logger.info(f"Cache atualizado para {key}")
    
    def invalidate_all(self):
        """Invalida todo o cache"""
        self._cache.clear()
        self._last_clear = datetime.now()
        logger.info("Cache invalidado completamente")
    
    def invalidate(self, key: str):
        """Invalida uma chave específica"""
        if key in self._cache:
            del self._cache[key]
            logger.info(f"Cache invalidado para {key}")

# ============================================
# TOKEN E AUTENTICAÇÃO
# ============================================
class GraphAuth:
    """Gerencia token de acesso à Microsoft Graph"""
    
    def __init__(self):
        self.access_token: Optional[str] = None
        self.token_expires: Optional[datetime] = None
        
    def get_token(self) -> str:
        """Obtém token de acesso (com cache automático)"""
        # Reusa token válido
        if self.access_token and self.token_expires and datetime.now() < self.token_expires:
            return self.access_token
            
        if not GraphConfig.is_configured():
            raise ValueError(
                "Variáveis de ambiente não configuradas. "
                "Verifique AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET e SHAREPOINT_SITE_URL"
            )
            
        url = f"https://login.microsoftonline.com/{GraphConfig.TENANT_ID}/oauth2/v2.0/token"
        
        data = {
            "grant_type": "client_credentials",
            "client_id": GraphConfig.CLIENT_ID,
            "client_secret": GraphConfig.CLIENT_SECRET,
            "scope": "https://graph.microsoft.com/.default"
        }
        
        try:
            logger.info("Obtendo novo token de acesso...")
            response = requests.post(url, data=data, timeout=30)
            response.raise_for_status()
            result = response.json()
            
            self.access_token = result["access_token"]
            expires_in = result.get("expires_in", 3600)
            self.token_expires = datetime.now() + timedelta(seconds=expires_in - 300)
            
            logger.info(f"Token obtido com sucesso. Expira em {expires_in}s")
            return self.access_token
            
        except requests.exceptions.RequestException as e:
            logger.error(f"Erro ao obter token: {str(e)}")
            raise ConnectionError(f"Erro ao obter token: {str(e)}")
        except KeyError as e:
            logger.error(f"Resposta inválida do Azure AD: {str(e)}")
            raise ValueError(f"Resposta inválida do Azure AD - verifique credenciais: {str(e)}")

# ============================================
# CLIENTE GRAPH
# ============================================
class GraphClient:
    """Cliente para chamadas à Microsoft Graph API"""
    
    def __init__(self):
        initialize_graph()
        self.auth = GraphAuth()
        self.cache = GraphCache()
        self.base_url = "https://graph.microsoft.com/v1.0"
        self._site_id: Optional[str] = None
        self._drive_id: Optional[str] = None
        self._file_id: Optional[str] = None
        
    def _headers(self) -> Dict[str, str]:
        """Headers padrão para requisições"""
        return {
            "Authorization": f"Bearer {self.auth.get_token()}",
            "Content-Type": "application/json"
        }
    
    def _get(self, endpoint: str, params: Optional[Dict] = None, retries: int = 3) -> Dict[str, Any]:
        """GET genérico na Graph API com retry"""
        url = f"{self.base_url}{endpoint}"
        
        for attempt in range(retries):
            try:
                response = requests.get(url, headers=self._headers(), params=params, timeout=30)
                response.raise_for_status()
                return response.json()
            except requests.exceptions.RequestException as e:
                logger.warning(f"Tentativa {attempt+1}/{retries} falhou: {str(e)}")
                if attempt == retries - 1:
                    raise
                time.sleep(2 ** attempt)
        
        raise RuntimeError(f"Falha após {retries} tentativas")
    
    # ============================================
    # SHAREPOINT / EXCEL
    # ============================================
    
    def get_site_id(self, site_url: str) -> str:
        """Obtém ID do site do SharePoint pela URL"""
        if self._site_id:
            return self._site_id
            
        if GraphConfig.is_onedrive_url():
            logger.info("Detectado OneDrive pessoal, usando método alternativo...")
            try:
                endpoint = "/me/drive"
                result = self._get(endpoint)
                self._site_id = "onedrive_personal"
                self._drive_id = result["id"]
                logger.info(f"OneDrive pessoal conectado: {self._drive_id}")
                return self._site_id
            except Exception as e:
                logger.warning(f"Falha no OneDrive pessoal: {str(e)}")
                parsed = urlparse(site_url)
                hostname = parsed.netloc
                site_path = parsed.path.strip("/")
                endpoint = f"/sites/{hostname}:/{site_path}"
                result = self._get(endpoint)
                self._site_id = result["id"]
                return self._site_id
        
        parsed = urlparse(site_url)
        hostname = parsed.netloc
        site_path = parsed.path.strip("/")
        
        endpoint = f"/sites/{hostname}:/{site_path}"
        try:
            result = self._get(endpoint)
            self._site_id = result["id"]
            logger.info(f"Site ID obtido: {self._site_id}")
            return self._site_id
        except Exception as e:
            logger.error(f"Erro ao obter site ID: {str(e)}")
            raise
    
    def get_drive_id(self, site_id: str) -> str:
        """Obtém ID do drive (biblioteca de documentos) do site"""
        if self._drive_id and site_id != "onedrive_personal":
            return self._drive_id
        
        if site_id == "onedrive_personal":
            return self._drive_id
            
        endpoint = f"/sites/{site_id}/drive"
        try:
            result = self._get(endpoint)
            self._drive_id = result["id"]
            logger.info(f"Drive ID obtido: {self._drive_id}")
            return self._drive_id
        except Exception as e:
            logger.error(f"Erro ao obter drive ID: {str(e)}")
            raise
    
    def get_file_id(self, drive_id: str, filename: str) -> str:
        """Busca arquivo pelo nome no drive"""
        if self._file_id:
            return self._file_id
        
        if drive_id == "onedrive_personal":
            try:
                endpoint = f"/me/drive/root/search(q='{filename}')"
                result = self._get(endpoint)
                
                if not result.get("value"):
                    raise FileNotFoundError(f"Arquivo '{filename}' não encontrado no OneDrive")
                
                for item in result["value"]:
                    if item.get("name") == filename:
                        self._file_id = item["id"]
                        logger.info(f"File ID obtido (OneDrive): {self._file_id}")
                        return self._file_id
                
                raise FileNotFoundError(f"Arquivo '{filename}' não encontrado")
            except Exception as e:
                logger.error(f"Erro no OneDrive: {str(e)}")
                raise
        
        endpoint = f"/drives/{drive_id}/root/search(q='{filename}')"
        try:
            result = self._get(endpoint)
            
            if not result.get("value"):
                raise FileNotFoundError(f"Arquivo '{filename}' não encontrado no SharePoint")
            
            for item in result["value"]:
                if item.get("name") == filename:
                    self._file_id = item["id"]
                    logger.info(f"File ID obtido: {self._file_id}")
                    return self._file_id
            
            raise FileNotFoundError(f"Arquivo '{filename}' não encontrado")
        except Exception as e:
            logger.error(f"Erro ao obter file ID: {str(e)}")
            raise
    
    def initialize_sharepoint(self) -> Tuple[str, str, str]:
        """Inicializa e retorna site_id, drive_id, file_id"""
        if not GraphConfig.is_configured():
            raise ValueError("Configuração do SharePoint incompleta")
        
        site_id = self.get_site_id(GraphConfig.SHAREPOINT_SITE_URL)
        drive_id = self.get_drive_id(site_id)
        file_id = self.get_file_id(drive_id, GraphConfig.EXCEL_FILENAME)
        
        return site_id, drive_id, file_id
    
    def get_excel_range(self, drive_id: str, file_id: str, worksheet_name: str, range_address: str = "A1:Z100") -> pd.DataFrame:
        """Lê dados de uma aba do Excel como DataFrame"""
        cache_key = f"{worksheet_name}_{range_address}"
        cached_data = self.cache.get(cache_key)
        if cached_data is not None:
            return cached_data
        
        try:
            if drive_id == "onedrive_personal":
                endpoint = f"/me/drive/items/{file_id}/workbook/worksheets('{worksheet_name}')/range(address='{range_address}')"
            else:
                endpoint = f"/drives/{drive_id}/items/{file_id}/workbook/worksheets('{worksheet_name}')/range(address='{range_address}')"
            
            result = self._get(endpoint)
            values = result.get("values", [])
            
            if not values or len(values) < 2:
                logger.warning(f"Nenhum dado encontrado na aba {worksheet_name}")
                return pd.DataFrame()
            
            headers = values[0]
            data = values[1:]
            
            df = pd.DataFrame(data, columns=headers)
            df = df.dropna(axis=1, how='all')
            
            for col in df.columns:
                if col and str(col).strip():
                    try:
                        df[col] = pd.to_numeric(df[col], errors='ignore')
                    except:
                        pass
            
            self.cache.set(cache_key, df)
            
            logger.info(f"Dados carregados da aba {worksheet_name}: {len(df)} linhas x {len(df.columns)} colunas")
            return df
            
        except Exception as e:
            logger.error(f"Erro ao ler Excel: {str(e)}")
            raise RuntimeError(f"Não foi possível ler a planilha: {str(e)}")
    
    def load_worksheet(self, worksheet_key: str) -> pd.DataFrame:
        """Carrega uma aba específica do Excel"""
        if worksheet_key not in GraphConfig.WORKSHEET_MAPPING:
            raise ValueError(f"Aba '{worksheet_key}' não mapeada")
        
        mapping = GraphConfig.WORKSHEET_MAPPING[worksheet_key]
        site_id, drive_id, file_id = self.initialize_sharepoint()
        
        return self.get_excel_range(
            drive_id,
            file_id,
            mapping["name"],
            mapping["range"]
        )
    
    def _converter_valor_monetario(self, valor) -> float:
        """
        Converte valores monetários formatados para float.
        CORRIGIDO: Suporta formato brasileiro (R$ 1.234,56)
        """
        if pd.isna(valor) or valor == '' or valor is None:
            return 0.0
        
        if isinstance(valor, (int, float)):
            return float(valor)
        
        if isinstance(valor, str):
            valor_limpo = valor.replace('R$', '').strip()
            valor_limpo = valor_limpo.replace('.', '').replace(',', '.')
            import re
            valor_limpo = re.sub(r'[^\d.]', '', valor_limpo)
            
            try:
                return float(valor_limpo) if valor_limpo else 0.0
            except ValueError:
                return 0.0
        
        return 0.0
    
    def _padronizar_colunas_controle(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Padroniza as colunas do DataFrame de controle.
        Detecta automaticamente as colunas e renomeia para o padrão esperado.
        """
        df = df.copy()
        df = df.dropna(how='all')
        
        if df.empty:
            return df
        
        meses = ['fev/26', 'mar/26', 'abr/26', 'mai/26', 'jun/26',
                 'jul/26', 'ago/26', 'set/26', 'out/26', 'nov/26', 'dez/26', 'jan/27']
        
        colunas_mapeadas = {}
        
        for idx, col in enumerate(df.columns):
            col_str = str(col).strip().lower()
            
            if idx == 0:
                colunas_mapeadas[col] = 'veiculo'
            elif idx == 1:
                colunas_mapeadas[col] = 'canal'
            elif idx == 2:
                colunas_mapeadas[col] = 'obs'
            elif any(mes.replace('/', '') in col_str.replace('/', '') for mes in meses):
                for mes in meses:
                    if mes.replace('/', '') in col_str.replace('/', ''):
                        colunas_mapeadas[col] = mes
                        break
            elif idx == len(df.columns) - 1:
                colunas_mapeadas[col] = 'total'
        
        df = df.rename(columns=colunas_mapeadas)
        
        colunas_manter = ['veiculo', 'canal', 'obs'] + meses + ['total']
        df = df[[col for col in colunas_manter if col in df.columns]]
        
        for mes in meses:
            if mes not in df.columns:
                df[mes] = 0.0
        
        for mes in meses:
            if mes in df.columns:
                try:
                    df[mes] = df[mes].apply(self._converter_valor_monetario)
                except Exception as e:
                    logger.warning(f"Erro ao converter mês {mes}: {str(e)}")
                    df[mes] = 0.0
        
        if 'total' not in df.columns:
            try:
                df['total'] = df[meses].sum(axis=1)
            except Exception as e:
                logger.warning(f"Erro ao calcular total: {str(e)}")
                df['total'] = 0.0
        
        logger.info(f"Colunas padronizadas: {list(df.columns)}")
        return df
    
    def get_big_numbers(self) -> Dict[str, float]:
        
        try:
            df_controle = self.load_worksheet("controle")
            
            if df_controle.empty:
                logger.warning("DataFrame de controle vazio")
                return {
                    "total_controle_aprovado": 0,
                    "total_controle_utilizado": 0,
                    "total_saldo_positivo": 0
                }
            
            df_controle = self._padronizar_colunas_controle(df_controle)
            
            # ===== DEBUG: Mostra todas as linhas =====
            logger.info("🔍 Procurando linhas de total...")
            for idx, row in df_controle.iterrows():
                if pd.notna(row['veiculo']):
                    logger.info(f"  Linha {idx}: {str(row['veiculo'])[:50]}...")
            
            big_numbers = {
                "total_controle_aprovado": 0,
                "total_controle_utilizado": 0,
                "total_saldo_positivo": 0
            }
            
            # ===== PALAVRAS-CHAVE MAIS AMPLAS =====
            keywords = {
                "aprovado": ["APROV", "APROVADO", "APROV.", "APROVADO", "TOTAL CONTROLE APROV", "TOTAL CONTROLE APROVADO"],
                "utilizado": ["UTILIZ", "UTILIZADO", "UTILIZ.", "UTILIZADO", "TOTAL CONTROLE UTILIZ", "TOTAL CONTROLE UTILIZADO"],
                "saldo": ["SALDO", "POSITIVO", "SALDO POSITIVO", "TOTAL SALDO POSITIVO"]
            }
            
            # ===== PERCORRE TODAS AS LINHAS =====
            for idx, row in df_controle.iterrows():
                primeira_col = str(row['veiculo']).strip().upper() if pd.notna(row['veiculo']) else ''
                
                # Verifica se contém "TOTAL" + palavra-chave
                if 'TOTAL' in primeira_col:
                    if any(kw in primeira_col for kw in keywords["aprovado"]):
                        big_numbers["total_controle_aprovado"] = self._converter_valor_monetario(row['total'])
                        logger.info(f"✅ Total Controle Aprovado: R$ {big_numbers['total_controle_aprovado']:,.2f}")
                        
                    elif any(kw in primeira_col for kw in keywords["utilizado"]):
                        big_numbers["total_controle_utilizado"] = self._converter_valor_monetario(row['total'])
                        logger.info(f"✅ Total Controle Utilizado: R$ {big_numbers['total_controle_utilizado']:,.2f}")
                        
                    elif any(kw in primeira_col for kw in keywords["saldo"]):
                        big_numbers["total_saldo_positivo"] = self._converter_valor_monetario(row['total'])
                        logger.info(f"✅ Total Saldo Positivo: R$ {big_numbers['total_saldo_positivo']:,.2f}")
            
            # ===== FALLBACK: Busca por posição (últimas 5 linhas) =====
            if big_numbers["total_controle_aprovado"] == 0:
                # Procura nas últimas 10 linhas por "APROV"
                for idx in range(len(df_controle) - 1, max(0, len(df_controle) - 10), -1):
                    row = df_controle.iloc[idx]
                    desc = str(row['veiculo']).strip().upper() if pd.notna(row['veiculo']) else ''
                    if 'APROV' in desc or 'APROVADO' in desc:
                        big_numbers["total_controle_aprovado"] = self._converter_valor_monetario(row['total'])
                        logger.info(f"✅ (Fallback posição) Total Controle Aprovado: R$ {big_numbers['total_controle_aprovado']:,.2f}")
                        break
            
            if big_numbers["total_controle_utilizado"] == 0:
                for idx in range(len(df_controle) - 1, max(0, len(df_controle) - 10), -1):
                    row = df_controle.iloc[idx]
                    desc = str(row['veiculo']).strip().upper() if pd.notna(row['veiculo']) else ''
                    if 'UTILIZ' in desc or 'UTILIZADO' in desc:
                        big_numbers["total_controle_utilizado"] = self._converter_valor_monetario(row['total'])
                        logger.info(f"✅ (Fallback posição) Total Controle Utilizado: R$ {big_numbers['total_controle_utilizado']:,.2f}")
                        break
            
            if big_numbers["total_saldo_positivo"] == 0:
                for idx in range(len(df_controle) - 1, max(0, len(df_controle) - 10), -1):
                    row = df_controle.iloc[idx]
                    desc = str(row['veiculo']).strip().upper() if pd.notna(row['veiculo']) else ''
                    if 'SALDO' in desc or 'POSITIVO' in desc:
                        big_numbers["total_saldo_positivo"] = self._converter_valor_monetario(row['total'])
                        logger.info(f"✅ (Fallback posição) Total Saldo Positivo: R$ {big_numbers['total_saldo_positivo']:,.2f}")
                        break
            
            return big_numbers
            
        except Exception as e:
            logger.error(f"Erro ao extrair big numbers: {str(e)}")
            return {
                "total_controle_aprovado": 0,
                "total_controle_utilizado": 0,
                "total_saldo_positivo": 0
            }
    
    def get_totais_mensais(self) -> Dict[str, Dict[str, float]]:
        try:
            df_controle = self.load_worksheet("controle")
            
            if df_controle.empty:
                return {}
            
            df_controle = self._padronizar_colunas_controle(df_controle)
            
            meses = ['fev/26', 'mar/26', 'abr/26', 'mai/26', 'jun/26',
                    'jul/26', 'ago/26', 'set/26', 'out/26', 'nov/26', 'dez/26', 'jan/27']
            
            mes_cols = [col for col in meses if col in df_controle.columns]
            
            if not mes_cols:
                logger.warning("Nenhuma coluna de mês encontrada")
                return {}
            
            keywords = {
                "aprovado": ["APROV", "APROVADO", "APROV.", "TOTAL CONTROLE APROV"],
                "utilizado": ["UTILIZ", "UTILIZADO", "UTILIZ.", "TOTAL CONTROLE UTILIZ"],
                "saldo": ["SALDO", "POSITIVO", "SALDO POSITIVO"]
            }
            
            totais = {
                "aprovado": {},
                "utilizado": {},
                "saldo": {}
            }
            
            for idx, row in df_controle.iterrows():
                primeira_col = str(row['veiculo']).strip().upper() if pd.notna(row['veiculo']) else ''
                
                if 'TOTAL' in primeira_col:
                    if any(kw in primeira_col for kw in keywords["aprovado"]):
                        for col in mes_cols:
                            if col in row.index:
                                totais["aprovado"][col] = self._converter_valor_monetario(row[col])
                        logger.info(f"✅ Totais mensais aprovados: {len(totais['aprovado'])} meses")
                        
                    elif any(kw in primeira_col for kw in keywords["utilizado"]):
                        for col in mes_cols:
                            if col in row.index:
                                totais["utilizado"][col] = self._converter_valor_monetario(row[col])
                        logger.info(f"✅ Totais mensais utilizados: {len(totais['utilizado'])} meses")
                        
                    elif any(kw in primeira_col for kw in keywords["saldo"]):
                        for col in mes_cols:
                            if col in row.index:
                                totais["saldo"][col] = self._converter_valor_monetario(row[col])
                        logger.info(f"✅ Totais mensais saldo: {len(totais['saldo'])} meses")
            
            return totais
            
        except Exception as e:
            logger.error(f"Erro ao extrair totais mensais: {str(e)}")
            return {}
    
    def load_dashboard_data(self) -> Dict[str, Any]:
        """Carrega todos os dados necessários para o dashboard"""
        site_id, drive_id, file_id = self.initialize_sharepoint()
        
        result = {
            "timestamp": datetime.now(),
            "source": "sharepoint",
            "file_name": GraphConfig.EXCEL_FILENAME
        }
        
        for key, mapping in GraphConfig.WORKSHEET_MAPPING.items():
            try:
                df = self.get_excel_range(
                    drive_id, 
                    file_id, 
                    mapping["name"], 
                    mapping["range"]
                )
                if key == "controle" and not df.empty:
                    df = self._padronizar_colunas_controle(df)
                result[key] = df
                logger.info(f"Aba '{key}' carregada com sucesso: {len(df)} linhas")
            except Exception as e:
                logger.error(f"Erro ao carregar aba '{key}': {str(e)}")
                result[key] = pd.DataFrame()
        
        return result

# ============================================
# FUNÇÃO DE VALIDAÇÃO E TESTE
# ============================================
def test_sharepoint_connection(quick: bool = False) -> Tuple[bool, str]:
    """
    Testa a conexão com o SharePoint.
    Se quick=True, faz apenas um teste rápido de autenticação.
    """
    try:
        if not GraphConfig.is_configured():
            return False, "Variáveis de ambiente não configuradas"
        
        # Teste rápido: apenas verifica se o token pode ser obtido
        if quick:
            try:
                client = GraphClient()
                token = client.auth.get_token()
                if token:
                    return True, "Autenticação OK"
                return False, "Falha na autenticação"
            except Exception as e:
                return False, f"Erro de autenticação: {str(e)}"
        
        # Teste completo: carrega a planilha
        client = GraphClient()
        site_id, drive_id, file_id = client.initialize_sharepoint()
        
        df = client.load_worksheet("controle")
        df = client._padronizar_colunas_controle(df)
        
        return True, f"Conexão OK! Linhas: {len(df)}"
        
    except requests.exceptions.HTTPError as e:
        if e.response.status_code == 403:
            return False, "🚫 Erro 403: O app não tem permissão para acessar este arquivo. Verifique as permissões no Azure AD."
        return False, f"Erro HTTP: {str(e)}"
    except Exception as e:
        return False, f"Erro: {str(e)}"

# ============================================
# FUNÇÕES PARA STREAMLIT (COM CACHE)
# ============================================
@st.cache_data(ttl=600)
def load_sharepoint_data(_client: GraphClient = None) -> Dict[str, Any]:
    """
    Função wrapper para cache do Streamlit.
    Carrega dados do SharePoint com cache.
    """
    initialize_graph()
    
    if _client is None:
        _client = GraphClient()
    
    try:
        data = _client.load_dashboard_data()
        data["_carregado_em"] = datetime.now().strftime("%H:%M:%S")
        return data
    except Exception as e:
        logger.error(f"Erro ao carregar dados do SharePoint: {str(e)}")
        return {
            "error": str(e),
            "timestamp": datetime.now(),
            "source": "error",
            "_carregado_em": datetime.now().strftime("%H:%M:%S")
        }

@st.cache_data(ttl=0)
def load_sharepoint_data_forced(_client: GraphClient = None) -> Dict[str, Any]:
    """
    Carrega dados do SharePoint FORÇANDO a recarga (ignora cache).
    """
    initialize_graph()
    invalidate_sharepoint_cache()
    
    if _client is None:
        _client = GraphClient()
    
    try:
        data = _client.load_dashboard_data()
        data["_carregado_em"] = datetime.now().strftime("%H:%M:%S")
        data["_forcado"] = True
        return data
    except Exception as e:
        logger.error(f"Erro ao carregar dados do SharePoint (forçado): {str(e)}")
        return {
            "error": str(e),
            "timestamp": datetime.now(),
            "source": "error",
            "_carregado_em": datetime.now().strftime("%H:%M:%S"),
            "_forcado": True
        }

@st.cache_data(ttl=300)
def get_sharepoint_big_numbers(_client: GraphClient = None) -> Dict[str, float]:
    """Obtém big numbers com cache do Streamlit"""
    initialize_graph()
    
    if _client is None:
        _client = GraphClient()
    
    try:
        return _client.get_big_numbers()
    except Exception as e:
        logger.error(f"Erro ao obter big numbers: {str(e)}")
        return {
            "total_controle_aprovado": 0,
            "total_controle_utilizado": 0,
            "total_saldo_positivo": 0
        }

@st.cache_data(ttl=300)
def get_sharepoint_totais_mensais(_client: GraphClient = None) -> Dict[str, Dict[str, float]]:
    """Obtém totais mensais com cache do Streamlit"""
    initialize_graph()
    
    if _client is None:
        _client = GraphClient()
    
    try:
        return _client.get_totais_mensais()
    except Exception as e:
        logger.error(f"Erro ao obter totais mensais: {str(e)}")
        return {}

def invalidate_sharepoint_cache():
    """Invalida todos os caches do SharePoint"""
    try:
        st.cache_data.clear()
        logger.info("Cache do Streamlit invalidado")
    except Exception as e:
        logger.warning(f"Erro ao limpar cache do Streamlit: {str(e)}")
    
    try:
        client = GraphClient()
        client.cache.invalidate_all()
        logger.info("Cache do GraphClient invalidado")
    except Exception as e:
        logger.warning(f"Erro ao limpar cache do GraphClient: {str(e)}")

# ============================================
# MAIN PARA TESTE DIRETO
# ============================================
if __name__ == "__main__":
    print("\n=== TESTE DE CONEXÃO SHAREPOINT ===\n")
    
    print("📋 Configurações carregadas:")
    print(f"  Tenant ID: {GraphConfig.TENANT_ID[:8]}...{GraphConfig.TENANT_ID[-4:] if len(GraphConfig.TENANT_ID) > 12 else '***'}")
    print(f"  Client ID: {GraphConfig.CLIENT_ID[:8]}...{GraphConfig.CLIENT_ID[-4:] if len(GraphConfig.CLIENT_ID) > 12 else '***'}")
    print(f"  Client Secret: {'✅ Configurado' if GraphConfig.CLIENT_SECRET else '❌ Não configurado'}")
    print(f"  Site URL: {GraphConfig.SHAREPOINT_SITE_URL if GraphConfig.SHAREPOINT_SITE_URL else '❌ Não configurado'}")
    print(f"  Tipo: {'OneDrive Pessoal' if GraphConfig.is_onedrive_url() else 'SharePoint'}")
    print(f"  Arquivo: {GraphConfig.EXCEL_FILENAME}")
    print(f"  Cache: {GraphConfig.CACHE_MINUTES} min")
    print()
    
    # Teste rápido
    print("🔍 Teste rápido de autenticação...")
    success, message = test_sharepoint_connection(quick=True)
    print(f"  Status: {'✅' if success else '❌'} {message}")
    
    if success:
        print("\n🚀 Teste completo: Carregando dados...")
        success, message = test_sharepoint_connection(quick=False)
        print(f"  Status: {'✅' if success else '❌'} {message}")
        
        if success:
            client = GraphClient()
            
            try:
                big_numbers = client.get_big_numbers()
                print(f"\n✅ Big Numbers carregados:")
                print(f"  Total Controle Aprovado: R$ {big_numbers['total_controle_aprovado']:,.2f}")
                print(f"  Total Controle Utilizado: R$ {big_numbers['total_controle_utilizado']:,.2f}")
                print(f"  Total Saldo Positivo: R$ {big_numbers['total_saldo_positivo']:,.2f}")
                
                totais = client.get_totais_mensais()
                if totais and totais.get('aprovado'):
                    print(f"\n📊 Totais mensais encontrados:")
                    print(f"  Aprovado: {len(totais['aprovado'])} meses")
                    print(f"  Utilizado: {len(totais['utilizado'])} meses")
                    print(f"  Saldo: {len(totais['saldo'])} meses")
                    
                    meses = list(totais['aprovado'].keys())[:3]
                    for mes in meses:
                        print(f"    {mes}: Aprovado R$ {totais['aprovado'].get(mes, 0):,.2f} | Utilizado R$ {totais['utilizado'].get(mes, 0):,.2f} | Saldo R$ {totais['saldo'].get(mes, 0):,.2f}")
                
                data = client.load_dashboard_data()
                print(f"\n📊 Abas carregadas:")
                for key, df in data.items():
                    if isinstance(df, pd.DataFrame):
                        print(f"  • {key}: {len(df)} linhas x {len(df.columns)} colunas")
                    elif key not in ['timestamp', 'source', 'file_name', '_carregado_em']:
                        print(f"  • {key}: {type(df).__name__}")
                
                print("\n✅ Teste concluído com sucesso!")
                
            except Exception as e:
                print(f"❌ Erro ao carregar dados: {e}")
                import traceback
                traceback.print_exc()
    else:
        if "403" in message:
            print("\n🔧 SOLUÇÃO PARA ERRO 403:")
            print("  1. No Azure Portal, vá para seu App Registration")
            print("  2. Clique em 'API Permissions'")
            print("  3. Adicione as permissões:")
            print("     • Microsoft Graph → Application Permissions → Files.Read.All")
            print("     • Microsoft Graph → Application Permissions → Sites.Read.All")
            print("  4. Clique em 'Grant admin consent'")
            print("  5. Aguarde alguns minutos e tente novamente")
        else:
            print(f"\n🔧 Erro detectado: {message}")
            print("  Verifique suas configurações e tente novamente.")