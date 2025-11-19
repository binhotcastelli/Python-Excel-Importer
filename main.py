import os
import sys
from data_import import import_from_excel
from reports import generate_complete_report

def main():
    # Configurações
    INPUT_DIR = "data/input"
    OUTPUT_DIR = "data/output"
    
    # Cria diretórios se não existirem
    os.makedirs(INPUT_DIR, exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    
    # Lista arquivos Excel disponíveis
    excel_files = [f for f in os.listdir(INPUT_DIR) if f.endswith(('.xlsx', '.xls'))]
    
    if not excel_files:
        print("Nenhum arquivo Excel encontrado na pasta data/input/")
        return
    
    print("Arquivos Excel encontrados:")
    for i, file in enumerate(excel_files, 1):
        print(f"{i}. {file}")
    
    # Seleciona arquivo
    try:
        choice = int(input("\nSelecione o número do arquivo: ")) - 1
        selected_file = excel_files[choice]
        file_path = os.path.join(INPUT_DIR, selected_file)
    except (ValueError, IndexError):
        print("Seleção inválida")
        return
    
    # Importa dados
    print(f"\nImportando {selected_file}...")
    dataframe, info = import_from_excel(file_path)
    
    if dataframe is not None:
        print(f"✅ Importação bem-sucedida!")
        print(f"📊 Registros: {info['total_registros']}")
        print(f"📈 Colunas: {info['total_colunas']}")
        
        # Gera relatórios
        print("\n📋 Gerando relatórios...")
        excel_report, text_report = generate_complete_report(dataframe, OUTPUT_DIR)
        
        print(f"✅ Relatório Excel: {excel_report}")
        print(f"✅ Sumário textual: {text_report}")
        
        # Mostra preview dos dados
        print("\n🔍 Preview dos dados (primeiras 5 linhas):")
        print(dataframe.head())
        
    else:
        print("❌ Falha na importação")

if __name__ == "__main__":
    main()