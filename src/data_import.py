import pandas as pd
import os
from datetime import datetime
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class ExcelDataImporter:
    def __init__(self, file_path):
        self.file_path = file_path
        self.df = None
    
    def load_excel(self, sheet_name=0, header=0):
        """Carrega arquivo Excel"""
        try:
            self.df = pd.read_excel(
                self.file_path, 
                sheet_name=sheet_name, 
                header=header,
                engine='openpyxl'
            )
            logger.info(f"Arquivo carregado: {len(self.df)} registros")
            return True
        except Exception as e:
            logger.error(f"Erro ao carregar arquivo: {e}")
            return False
    
    def clean_data(self):
        """Limpeza básica dos dados"""
        if self.df is None:
            return False
        
        # Remove colunas completamente vazias
        self.df.dropna(axis=1, how='all', inplace=True)
        
        # Remove linhas completamente vazias
        self.df.dropna(how='all', inplace=True)
        
        # Preenche valores numéricos faltantes com 0
        numeric_cols = self.df.select_dtypes(include=['number']).columns
        self.df[numeric_cols] = self.df[numeric_cols].fillna(0)
        
        # Remove espaços extras em strings
        string_cols = self.df.select_dtypes(include=['object']).columns
        self.df[string_cols] = self.df[string_cols].apply(
            lambda x: x.str.strip() if x.dtype == "object" else x
        )
        
        logger.info("Dados limpos com sucesso")
        return True
    
    def get_basic_info(self):
        """Retorna informações básicas do dataset"""
        if self.df is None:
            return {}
        
        info = {
            'total_registros': len(self.df),
            'total_colunas': len(self.df.columns),
            'colunas': list(self.df.columns),
            'tipos_dados': self.df.dtypes.to_dict(),
            'registros_faltantes': self.df.isnull().sum().to_dict()
        }
        return info

def import_from_excel(file_path, clean=True):
    """Função principal de importação"""
    importer = ExcelDataImporter(file_path)
    
    if importer.load_excel():
        if clean:
            importer.clean_data()
        
        info = importer.get_basic_info()
        logger.info(f"Importação concluída: {info['total_registros']} registros")
        
        return importer.df, info
    
    return None, {}