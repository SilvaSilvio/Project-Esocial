import pandas as pd
import numpy as np
from datetime import datetime
import re
import os

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
        
        # Separa as colunas fixas e as colunas a serem transformadas
        colunas_fixas = df.columns[:categoria_idx + 1].tolist()
        colunas_transformar = df.columns[categoria_idx + 1:].tolist()
        
        # Lista para armazenar as linhas transformadas
        linhas_transformadas = []
        
        # Para cada linha do DataFrame original
        for _, row in df.iterrows():
            valores_fixos = row[colunas_fixas].to_dict()
            
            # Para cada coluna que deve ser transformada em linha
            for coluna in colunas_transformar:
                valor = row[coluna]
                # Só inclui se o valor não for nulo, vazio ou zero
                try:
                    if pd.notna(valor) and str(valor).strip() and str(valor).strip() != '0':
                        nova_linha = valores_fixos.copy()
                        codigo, descricao = extrair_codigo_descricao(coluna)
                        nova_linha['Código'] = codigo
                        nova_linha['Descrição'] = descricao
                        nova_linha['Descrição Completa'] = coluna
                        
                        # Converte o valor para o formato correto
                        try:
                            valor_str = str(valor).replace('.', '').replace(',', '.')
                            valor_float = float(valor_str)
                            # Se o valor for muito grande, assume que está em centavos
                            if valor_float > 100000:
                                valor_float = valor_float / 100
                            nova_linha['Valor'] = round(valor_float, 2)
                        except:
                            nova_linha['Valor'] = 0.0
                        
                        linhas_transformadas.append(nova_linha)
                except:
                    continue
        
        if linhas_transformadas:
            result_df = pd.DataFrame(linhas_transformadas)
            colunas_finais = ['MATRICULA', 'Código', 'Descrição', 'Descrição Completa', 'Valor']
            
            # Garante que todas as colunas necessárias existam
            for col in colunas_finais:
                if col not in result_df.columns:
                    result_df[col] = ''
            
            result_df = result_df[colunas_finais]
            result_df = result_df.sort_values(by=['MATRICULA', 'Código'], na_position='last')
            
            # Converte a coluna MATRICULA para string
            result_df['MATRICULA'] = result_df['MATRICULA'].astype(str)
            
            return result_df
        else:
            return pd.DataFrame(columns=['MATRICULA', 'Código', 'Descrição', 'Descrição Completa', 'Valor'])
    
    except Exception as e:
        print(f"Erro ao transformar dados: {str(e)}")
        return None

def group_and_sum(df):
    """
    Agrupa os registros iguais e soma os valores
    """
    try:
        print("\nIniciando agrupamento...")
        
        # Garante que a coluna Valor está em formato numérico
        df['Valor'] = pd.to_numeric(df['Valor'].astype(str).str.replace('.', '').str.replace(',', '.'), errors='coerce')
        
        # Define as colunas para agrupamento (todas exceto Valor)
        group_columns = [col for col in df.columns if col != 'Valor']
        
        # Mostra exemplo antes do agrupamento
        print("\nExemplo antes do agrupamento:")
        print(df[['Código', 'Descrição', 'Valor']].head())
        print(f"Total de registros: {len(df)}")
        
        # Agrupa e soma os valores
        df_grouped = df.groupby(group_columns, as_index=False).agg({
            'Valor': 'sum'
        })
        
        # Mostra exemplo após o agrupamento
        print("\nExemplo após agrupamento:")
        print(df_grouped[['Código', 'Descrição', 'Valor']].head())
        print(f"Total de registros após agrupamento: {len(df_grouped)}")
        
        # Ordena o resultado
        df_grouped = df_grouped.sort_values(by=['Código', 'Valor'], na_position='last')
        
        # Formata o valor para duas casas decimais
        df_grouped['Valor'] = df_grouped['Valor'].round(2)
        
        return df_grouped
        
    except Exception as e:
        print(f"\nERRO no agrupamento: {str(e)}")
        raise e

def process_esocial(file_path):
    try:
        # Verifica se o arquivo existe
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Arquivo não encontrado: {file_path}")
            
        print(f"Processando arquivo: {file_path}")
        
        # Lê o arquivo Excel
        df = pd.read_excel(file_path)
        print(f"\nArquivo lido com sucesso. Total de linhas: {len(df)}")
        
        # Transforma os dados
        print("\nTransformando dados...")
        df_transformed = transform_data(df)
        if df_transformed is None:
            raise ValueError("Erro na transformação dos dados")
        
        print(f"Dados transformados com sucesso. Total de linhas: {len(df_transformed)}")
        
        # Lista de descrições a serem removidas
        descricoes_remover = [
            'CPF DEPENDENTE',
            'IR COMPETENCIA - CPF DEPENDENTE',
            'CNPJ OPERADORA SAUDE'
        ]
        
        # Remove as linhas com as descrições especificadas
        print("\nRemovendo descrições específicas...")
        total_antes = len(df_transformed)
        for desc in descricoes_remover:
            df_transformed = df_transformed[~df_transformed['Descrição Completa'].str.contains(desc, na=False, case=False)]
        print(f"Linhas removidas: {total_antes - len(df_transformed)}")
        
        # Agrupa os dados
        print("\nAgrupando dados...")
        grouped = df_transformed.groupby(['MATRICULA', 'Código', 'Descrição Completa'])['Valor'].sum().reset_index()
        print(f"Dados agrupados. Total de linhas: {len(grouped)}")
        
        # Ordena os resultados
        grouped = grouped.sort_values(['MATRICULA', 'Código'])
        
        # Formata os valores para duas casas decimais
        grouped['Valor'] = grouped['Valor'].round(2)
        
        # Gera o nome do arquivo de saída
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"esocial_processado_{timestamp}.xlsx"
        
        # Salva o resultado em Excel
        print(f"\nSalvando resultado em {output_file}...")
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            grouped.to_excel(writer, index=False, sheet_name='Dados')
            
            # Formata a coluna de valor
            workbook = writer.book
            worksheet = writer.sheets['Dados']
            
            # Encontra a coluna de valor
            valor_col = None
            for idx, col in enumerate(grouped.columns):
                if col == 'Valor':
                    valor_col = idx + 1  # Converte para índice baseado em 1
                    break
            
            if valor_col is not None:
                # Aplica o formato de número com duas casas decimais
                for row in range(2, len(grouped) + 2):  # +2 porque o Excel começa em 1 e tem o cabeçalho
                    cell = worksheet.cell(row=row, column=valor_col)
                    cell.number_format = '#,##0.00'
        
        print(f"Arquivo processado e salvo como: {output_file}")
        return grouped
        
    except Exception as e:
        print(f"\nErro ao processar arquivo: {str(e)}")
        import traceback
        print("\nDetalhes do erro:")
        print(traceback.format_exc())
        return None

if __name__ == "__main__":
    import sys
    
    if len(sys.argv) != 2:
        print("Uso: python process_esocial.py <arquivo_excel>")
    else:
        file_path = sys.argv[1]
        process_esocial(file_path) 