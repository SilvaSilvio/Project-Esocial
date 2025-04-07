import pandas as pd
import os
import re
from datetime import datetime

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
    Transforma o DataFrame para o formato desejado, convertendo colunas em linhas
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
                    # Converte o valor para string e trata como valor monetário
                    valor_str = str(valor).strip()
                    if valor_str and valor_str.lower() != 'nan':
                        # Verifica se o valor é numérico
                        if isinstance(valor, (int, float)) or valor_str.replace(',', '').replace('.', '').isdigit():
                            # Se for numérico, converte para float
                            if isinstance(valor, (int, float)):
                                # Se for um valor inteiro maior que 100, divide por 100
                                if isinstance(valor, int) and valor > 100:
                                    valor_numerico = float(valor) / 100
                                else:
                                    valor_numerico = float(valor)
                            else:
                                # Trata como valor monetário brasileiro
                                if ',' in valor_str and '.' in valor_str:
                                    # Formato brasileiro: 1.234,56
                                    valor_numerico = float(valor_str.replace('.', '').replace(',', '.'))
                                elif ',' in valor_str:
                                    # Apenas vírgula: 1234,56
                                    valor_numerico = float(valor_str.replace(',', '.'))
                                elif '.' in valor_str:
                                    # Apenas ponto: 1234.56
                                    valor_numerico = float(valor_str)
                                else:
                                    # Apenas números: 1234
                                    valor_numerico = float(valor_str) / 100
                            
                            if valor_numerico != 0:
                                nova_linha = valores_fixos.copy()
                                codigo, descricao = extrair_codigo_descricao(coluna)
                                nova_linha['Código'] = codigo
                                nova_linha['Descrição'] = descricao
                                nova_linha['Descrição Completa'] = coluna
                                nova_linha['Valor'] = valor_numerico
                                linhas_transformadas.append(nova_linha)
                except (ValueError, TypeError):
                    continue
        
        if linhas_transformadas:
            result_df = pd.DataFrame(linhas_transformadas)
            colunas_finais = colunas_fixas + ['Código', 'Descrição', 'Descrição Completa', 'Valor']
            result_df = result_df[colunas_finais]
            result_df = result_df.sort_values(by=['Código', 'Valor'], na_position='last')
            return result_df
        else:
            return pd.DataFrame(columns=colunas_fixas + ['Código', 'Descrição', 'Descrição Completa', 'Valor'])
    
    except Exception as e:
        print(f"Erro ao transformar dados: {str(e)}")
        return None

def process_esocial(input_file):
    """
    Processa o arquivo eSocial
    """
    try:
        print(f"\nProcessando arquivo eSocial: {input_file}")
        
        # Lê o arquivo Excel
        df_esocial = pd.read_excel(input_file)
        print("Arquivo lido com sucesso")
        
        # Transforma os dados
        df_transformed = transform_data(df_esocial)
        if df_transformed is None:
            raise ValueError("Erro na transformação dos dados")
        
        print("\nDados transformados com sucesso")
        
        # Gera o nome do arquivo de saída com timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"esocial_processado_{timestamp}.xlsx"
        
        # Salva o resultado em Excel
        print(f"\nSalvando resultado em {output_file}...")
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df_transformed.to_excel(writer, index=False, sheet_name='Dados')
            
            # Formata a planilha
            workbook = writer.book
            worksheet = writer.sheets['Dados']
            
            # Ajusta a largura das colunas
            for idx, col in enumerate(df_transformed.columns):
                max_length = max(
                    df_transformed[col].astype(str).apply(len).max(),
                    len(str(col))
                )
                worksheet.column_dimensions[chr(65 + idx)].width = max_length + 2
                
                # Formata a coluna de valor com duas casas decimais
                if col == 'Valor':
                    for cell in worksheet[chr(65 + idx)][1:]:
                        cell.number_format = '#,##0.00'
        
        print(f"Arquivo processado e salvo como: {output_file}")
        return df_transformed
        
    except Exception as e:
        print(f"\nErro ao processar arquivo: {str(e)}")
        return None

def main():
    if len(sys.argv) != 2:
        print("Uso: python process_esocial.py <arquivo_esocial>")
        print("Exemplo: python process_esocial.py esocial.xlsx")
        return
    
    file_path = sys.argv[1]
    if not os.path.exists(file_path):
        print(f"Erro: Arquivo não encontrado: {file_path}")
        return
    
    process_esocial(file_path)

if __name__ == "__main__":
    import sys
    main() 