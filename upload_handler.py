# ============================================
# upload_handler.py - Módulo para upload de arquivos
# ============================================
import streamlit as st
import pandas as pd
from pathlib import Path
import tempfile
import os

class UploadHandler:
    """Gerencia upload e processamento de arquivos"""
    
    @staticmethod
    def processar_upload(arquivo):
        """Processa arquivo enviado pelo usuário"""
        
        if arquivo is None:
            return None
            
        # Verificar extensão
        extensao = Path(arquivo.name).suffix.lower()
        if extensao not in ['.xlsx', '.xls', '.csv']:
            st.error("Formato não suportado. Use .xlsx, .xls ou .csv")
            return None
        
        # Salvar arquivo temporário
        with tempfile.NamedTemporaryFile(delete=False, suffix=extensao) as tmp:
            tmp.write(arquivo.getvalue())
            tmp_path = tmp.name
        
        try:
            # Carregar conforme extensão
            if extensao == '.csv':
                df = pd.read_csv(tmp_path, sep=';', decimal=',')
            else:
                df = pd.read_excel(tmp_path, sheet_name=None)
            
            return df
            
        except Exception as e:
            st.error(f"Erro ao processar arquivo: {str(e)}")
            return None
        finally:
            # Limpar arquivo temporário
            if os.path.exists(tmp_path):
                os.unlink(tmp_path)
    
    @staticmethod
    def validar_estrutura(dados):
        """Valida se o arquivo tem a estrutura esperada"""
        
        if isinstance(dados, dict):  # Múltiplas abas
            abas_necessarias = ['Geral', 'Geral Controle (2) nova']
            for aba in abas_necessarias:
                if aba not in dados:
                    return False, f"ABA '{aba}' não encontrada"
            
            # Verificar estrutura da aba Controle
            df_controle = dados['Geral Controle (2) nova']
            colunas_necessarias = ['veiculo', 'canal', 'obs']
            
            # Tentar detectar colunas
            colunas = df_controle.columns.tolist()
            if 'veiculo' not in colunas and 'Veículo' in colunas:
                df_controle.columns = df_controle.columns.str.lower()
            
            return True, "Estrutura válida"
        
        return False, "Formato de dados inválido"