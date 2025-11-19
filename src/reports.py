import pandas as pd
import matplotlib.pyplot as plt
import os
from datetime import datetime

class ReportGenerator:
    def __init__(self, dataframe):
        self.df = dataframe
        self.report_data = {}
    
    def generate_summary_stats(self):
        """Estatísticas sumarizadas"""
        stats = {
            'descricao_numericas': self.df.describe().to_dict(),
            'contagem_categorias': {},
            'correlacoes': self.df.select_dtypes(include=['number']).corr().to_dict()
        }
        
        # Contagem para colunas categóricas
        categorical_cols = self.df.select_dtypes(include=['object']).columns
        for col in categorical_cols:
            stats['contagem_categorias'][col] = self.df[col].value_counts().to_dict()
        
        self.report_data['estatisticas'] = stats
        return stats
    
    def create_excel_report(self, output_path):
        """Cria relatório completo em Excel"""
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            # Dados originais
            self.df.to_excel(writer, sheet_name='Dados_Originais', index=False)
            
            # Estatísticas
            stats_df = pd.DataFrame(self.df.describe())
            stats_df.to_excel(writer, sheet_name='Estatisticas')
            
            # Análise de valores únicos
            unique_analysis = pd.DataFrame({
                'Coluna': self.df.columns,
                'Tipo_Dado': self.df.dtypes.values,
                'Valores_Unicos': [self.df[col].nunique() for col in self.df.columns],
                'Valores_Nulos': self.df.isnull().sum().values
            })
            unique_analysis.to_excel(writer, sheet_name='Analise_Colunas', index=False)
    
    def create_summary_report(self):
        """Cria relatório sumarizado em texto"""
        summary = f"""
        RELATÓRIO DE ANÁLISE DE DADOS
        =============================
        Data: {datetime.now().strftime('%d/%m/%Y %H:%M')}
        Total de Registros: {len(self.df):,}
        Total de Colunas: {len(self.df.columns)}
        
        COLUNAS E TIPOS:
        {self.df.dtypes.to_string()}
        
        ESTATÍSTICAS BÁSICAS:
        {self.df.describe().to_string()}
        """
        return summary

def generate_complete_report(dataframe, output_dir):
    """Gera relatório completo"""
    generator = ReportGenerator(dataframe)
    
    # Gera estatísticas
    generator.generate_summary_stats()
    
    # Cria arquivo Excel
    excel_path = os.path.join(output_dir, f"relatorio_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx")
    generator.create_excel_report(excel_path)
    
    # Cria relatório textual
    text_report = generator.create_summary_report()
    
    # Salva relatório textual
    text_path = os.path.join(output_dir, f"sumario_{datetime.now().strftime('%Y%m%d_%H%M')}.txt")
    with open(text_path, 'w', encoding='utf-8') as f:
        f.write(text_report)
    
    return excel_path, text_path