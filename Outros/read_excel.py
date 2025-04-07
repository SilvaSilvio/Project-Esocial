import pandas as pd
import re
from datetime import datetime

def read_excel_file(file_path):
    """
    Lê um arquivo Excel e retorna um dicionário com todos os DataFrames
    """
    try:
        # Lê todas as planilhas do arquivo Excel
        excel_file = pd.ExcelFile(file_path)
        
        # Cria um dicionário para armazenar os DataFrames
        dataframes = {}
        
        # Lê cada planilha e armazena no dicionário
        for sheet_name in excel_file.sheet_names:
            dataframes[sheet_name] = pd.read_excel(file_path, sheet_name=sheet_name)
            
        return dataframes
    except Exception as e:
        print(f"Erro ao ler o arquivo: {str(e)}")
        return None

def extrair_codigo_descricao(texto):
    """
    Extrai o código e a descrição de uma string
    Exemplo: 'TIPO-11: RENDIMENTO TRIBUTAVEL' -> ('11', 'RENDIMENTO TRIBUTAVEL')
    """
    padrao = r'TIPO-(\d+):\s*(.+)'
    match = re.search(padrao, str(texto))
    
    if match:
        return match.group(1), match.group(2).strip()
    return '', texto.strip()

def transform_data(df):
    """
    Transforma o DataFrame para o formato desejado
    """
    try:
        # Encontra o índice da coluna 'CATEGORIA DO TRABALHADOR'
        categoria_idx = None
        for idx, col in enumerate(df.columns):
            if 'CATEGORIA DO TRABALHADOR' in str(col).upper():
                categoria_idx = idx
                break
        
        if categoria_idx is None:
            raise ValueError("Coluna 'CATEGORIA DO TRABALHADOR' não encontrada")
        
        # Separa as colunas fixas (até CATEGORIA DO TRABALHADOR) e as colunas a serem transformadas
        colunas_fixas = df.columns[:categoria_idx + 1].tolist()
        colunas_transformar = df.columns[categoria_idx + 1:].tolist()
        
        # Lista para armazenar as linhas transformadas
        linhas_transformadas = []
        
        # Para cada linha do DataFrame original
        for _, row in df.iterrows():
            # Valores fixos que serão repetidos
            valores_fixos = row[colunas_fixas].to_dict()
            
            # Para cada coluna que deve ser transformada em linha
            for coluna in colunas_transformar:
                valor = row[coluna]
                # Só inclui se o valor não for nulo, vazio ou zero
                try:
                    valor_str = str(valor).replace(',', '.')
                    if valor_str.strip() and valor_str.strip() != 'nan':
                        valor_numerico = float(valor_str)
                        if valor_numerico != 0:
                            nova_linha = valores_fixos.copy()
                            # Extrai código e descrição
                            codigo, descricao = extrair_codigo_descricao(coluna)
                            nova_linha['Código'] = codigo
                            nova_linha['Descrição'] = descricao
                            nova_linha['Descrição Completa'] = coluna  # Adiciona a descrição completa original
                            nova_linha['Valor'] = valor_numerico
                            linhas_transformadas.append(nova_linha)
                except (ValueError, TypeError):
                    # Se não for possível converter para número, ignora o valor
                    continue
        
        # Cria o novo DataFrame com as linhas transformadas
        if linhas_transformadas:
            result_df = pd.DataFrame(linhas_transformadas)
            # Organiza as colunas: primeiro as fixas, depois Código, Descrição, Descrição Completa e Valor
            colunas_finais = colunas_fixas + ['Código', 'Descrição', 'Descrição Completa', 'Valor']
            result_df = result_df[colunas_finais]
            
            # Ordena o DataFrame por Código (se existir) e Valor
            result_df = result_df.sort_values(by=['Código', 'Valor'], na_position='last')
            
            return result_df
        else:
            return pd.DataFrame(columns=colunas_fixas + ['Código', 'Descrição', 'Descrição Completa', 'Valor'])
    
    except Exception as e:
        print(f"Erro ao transformar dados: {str(e)}")
        return None

def list_sheets(file_path):
    """
    Lista todas as planilhas disponíveis no arquivo Excel
    """
    try:
        excel_file = pd.ExcelFile(file_path)
        return excel_file.sheet_names
    except Exception as e:
        print(f"Erro ao listar planilhas: {str(e)}")
        return None

def get_sheet_data(file_path, sheet_name):
    """
    Obtém os dados de uma planilha específica e aplica a transformação
    """
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        transformed_df = transform_data(df)
        return transformed_df
    except Exception as e:
        print(f"Erro ao obter dados da planilha: {str(e)}")
        return None

# Exemplo de uso
if __name__ == "__main__":
    file_path = "ANDREANI_dirf x esocial 1.xlsx"
    
    # Gera um timestamp para o nome do arquivo
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # Lista todas as planilhas
    sheets = list_sheets(file_path)
    if sheets:
        print("Planilhas disponíveis:", sheets)
        
        # Para cada planilha, mostra os dados transformados
        for sheet_name in sheets:
            df = get_sheet_data(file_path, sheet_name)
            if df is not None:
                print(f"\nDados transformados da planilha '{sheet_name}':")
                print("\nPrimeiras 5 linhas:")
                print(df.head())
                
                # Salva o resultado em um novo arquivo Excel com timestamp
                output_file = f"transformado_{timestamp}.xlsx"
                df.to_excel(output_file, index=False)
                print(f"\nArquivo salvo como: {output_file}") 