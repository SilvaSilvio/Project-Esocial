import pandas as pd
from datetime import datetime
import os
import sys
from process_esocial import process_esocial_file
from process_lg_depara import process_lg_depara

def main():
    try:
        # Verifica se os argumentos foram fornecidos
        if len(sys.argv) < 4:
            print("Uso: python main.py <arquivo_esocial> <arquivo_ficha_financeira> <arquivo_depara>")
            print("Exemplo: python main.py ANDREANI_dirf_x_esocial_1.xlsx ficha_financeira.csv depara.csv")
            return
        
        # Arquivos de entrada
        esocial_file = sys.argv[1]
        lg_file = sys.argv[2]
        depara_file = sys.argv[3]
        
        # Verifica se os arquivos existem
        for file_path in [esocial_file, lg_file, depara_file]:
            if not os.path.exists(file_path):
                print(f"Erro: Arquivo não encontrado: {file_path}")
                return
        
        print("=== Iniciando processamento ===")
        print(f"Arquivo eSocial: {esocial_file}")
        print(f"Arquivo Ficha Financeira: {lg_file}")
        print(f"Arquivo De/Para: {depara_file}")
        
        # Processa o arquivo eSocial
        print("\n1. Processando arquivo eSocial...")
        df_esocial = process_esocial_file(esocial_file)
        if df_esocial is None:
            raise ValueError("Erro ao processar arquivo eSocial")
        
        # Processa os arquivos LG e Depara
        print("\n2. Processando arquivos LG e Depara...")
        df_lg_depara = process_lg_depara(lg_file, depara_file)
        if df_lg_depara is None:
            raise ValueError("Erro ao processar arquivos LG e Depara")
        
        # Gera o nome do arquivo de saída com timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"resultado_comparacao_{timestamp}.xlsx"
        
        # Salva os resultados em diferentes abas do mesmo arquivo Excel
        print(f"\nSalvando resultados em {output_file}...")
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df_esocial.to_excel(writer, sheet_name='eSocial', index=False)
            df_lg_depara.to_excel(writer, sheet_name='LG_Depara', index=False)
        
        print("\nProcessamento concluído com sucesso!")
        print(f"Arquivo de saída: {output_file}")
        
    except Exception as e:
        print(f"\nErro durante o processamento: {str(e)}")

if __name__ == "__main__":
    main() 