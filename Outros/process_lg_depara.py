import pandas as pd
from datetime import datetime
import os
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

def encontrar_coluna(nome_coluna, lista_colunas):
    """
    Encontra a coluna correta independente da codificação
    """
    if nome_coluna in lista_colunas:
        return nome_coluna
    # Se não encontrou, procura por similaridade
    for col in lista_colunas:
        if col.replace('Ó', 'O').replace('Ã', 'A') == nome_coluna.replace('Ó', 'O').replace('Ã', 'A'):
            return col
    return None

def read_file(file_path):
    """
    Lê um arquivo CSV ou Excel baseado na extensão
    """
    try:
        file_extension = os.path.splitext(file_path)[1].lower()
        
        if file_extension == '.csv':
            try:
                # Primeira tentativa: ler o arquivo em modo binário e decodificar com cp1252
                with open(file_path, 'rb') as file:
                    raw_data = file.read()
                    # Remove BOM se existir
                    if raw_data.startswith(b'\xef\xbb\xbf'):
                        raw_data = raw_data[3:]
                    # Decodifica com cp1252 (Windows-1252)
                    content = raw_data.decode('cp1252')
                    # Remove caracteres especiais e quebras de linha extras
                    content = content.replace('\r', '').replace('\n\n', '\n')
                    # Usa StringIO para criar um arquivo em memória
                    from io import StringIO
                    df = pd.read_csv(StringIO(content), sep=';', low_memory=False)
                    print("Arquivo lido com sucesso usando cp1252")
                    return df
            except Exception as e:
                print(f"Falha na primeira tentativa: {str(e)}")
                
                # Segunda tentativa: tentar outros encodings comuns
                encodings = ['iso-8859-1', 'latin1', 'utf-8-sig', 'utf-8', 'utf-16']
                
                for encoding in encodings:
                    try:
                        print(f"Tentando ler com encoding: {encoding}")
                        df = pd.read_csv(file_path, sep=';', encoding=encoding, low_memory=False)
                        print(f"Arquivo lido com sucesso usando encoding: {encoding}")
                        return df
                    except Exception as e:
                        print(f"Falha com encoding {encoding}: {str(e)}")
                        continue
                
                raise ValueError("Não foi possível ler o arquivo com nenhum dos encodings disponíveis")
            
        elif file_extension in ['.xlsx', '.xls']:
            df = pd.read_excel(file_path)
            
            # Converte colunas que parecem ser códigos para string
            for col in df.columns:
                if 'CÓDIGO' in col.upper() or 'COD' in col.upper() or 'ID' in col.upper():
                    df[col] = df[col].astype(str).str.strip()
            
            return df
        else:
            raise ValueError(f"Formato de arquivo não suportado: {file_extension}")
    except Exception as e:
        raise Exception(f"Erro ao ler o arquivo {file_path}: {str(e)}")

def convert_valor_to_numeric(df, valor_col):
    """
    Converte a coluna VALOR para numérico, preservando valores específicos
    """
    # Cria uma cópia da coluna original para referência
    df['VALOR_ORIGINAL'] = df[valor_col].copy()
    
    # Converte para string e faz as limpezas necessárias
    df[valor_col] = df[valor_col].astype(str).str.strip()
    
    # Função para converter valores no formato brasileiro (1.234,56)
    def convert_br_number(x):
        try:
            if isinstance(x, (int, float)):
                return float(x)
            if pd.isna(x) or x == '':
                return 0.0
            # Remove espaços e caracteres não numéricos, exceto . e ,
            x = ''.join(c for c in str(x) if c.isdigit() or c in ',.').strip()
            if ',' in x and '.' in x:
                # Formato brasileiro: 1.234,56
                return float(x.replace('.', '').replace(',', '.'))
            elif ',' in x:
                # Apenas vírgula: 1234,56
                return float(x.replace(',', '.'))
            elif '.' in x:
                # Apenas ponto: 1234.56
                return float(x)
            else:
                # Apenas números: 1234
                return float(x)
        except:
            return None

    # Aplica a conversão
    df[valor_col] = df[valor_col].apply(convert_br_number)
    
    # Verifica valores que foram convertidos para NaN
    nan_mask = df[valor_col].isna()
    if nan_mask.any():
        print(f"\nValores não convertidos para numérico: {df.loc[nan_mask, 'VALOR_ORIGINAL'].unique()}")
        
        # Tenta recuperar os valores originais que não foram convertidos
        for idx in df[nan_mask].index:
            valor_original = str(df.at[idx, 'VALOR_ORIGINAL'])
            print(f"Tentando converter valor original: {valor_original}")
            try:
                if '.' in valor_original and ',' in valor_original:
                    valor_convertido = float(valor_original.replace('.', '').replace(',', '.'))
                    df.at[idx, valor_col] = valor_convertido
                    print(f"Valor corrigido: {valor_original} -> {valor_convertido}")
            except:
                print(f"Não foi possível converter o valor: {valor_original}")
    
    # Preenche os valores NaN restantes com 0
    df[valor_col] = df[valor_col].fillna(0)
    
    return df

def process_lg_depara(lg_file, depara_file):
    """
    Processa os arquivos LG e Depara, concatenando-os e mapeando os códigos do eSocial
    """
    try:
        print(f"\nProcessando arquivos:")
        print(f"LG: {lg_file}")
        print(f"Depara: {depara_file}")
        
        # Verifica se os arquivos existem
        if not os.path.exists(lg_file):
            raise FileNotFoundError(f"Arquivo LG não encontrado: {lg_file}")
        if not os.path.exists(depara_file):
            raise FileNotFoundError(f"Arquivo Depara não encontrado: {depara_file}")
        
        # Lê os arquivos
        print("\nLendo arquivo LG...")
        df_lg = read_file(lg_file)
        
        # Exibe os dados da Ficha Financeira no console
        print("\n=== DADOS DA FICHA FINANCEIRA ===")
        print(f"Total de registros: {len(df_lg)}")
        print("\nPrimeiras 5 linhas da Ficha Financeira:")
        print(df_lg.head().to_string())
        
        # Procura pelo evento 5 (salário)
        cod_evento_col = encontrar_coluna('CÓDIGO EVENTO', df_lg.columns)
        valor_col = encontrar_coluna('VALOR', df_lg.columns)
        
        if cod_evento_col and valor_col:
            print("\n=== EVENTO 5 (SALÁRIO) ===")
            df_lg[cod_evento_col] = df_lg[cod_evento_col].astype(str).str.strip()
            evento5 = df_lg[df_lg[cod_evento_col] == '5']
            
            if not evento5.empty:
                print(f"Registros encontrados: {len(evento5)}")
                print("\nValores do evento 5 antes da conversão:")
                print(evento5[[cod_evento_col, valor_col, 'NOME']].to_string())
                
                # Converte a coluna VALOR para numérico
                df_lg = convert_valor_to_numeric(df_lg, valor_col)
                
                # Verifica os valores após a conversão
                evento5_apos_conversao = df_lg[df_lg[cod_evento_col] == '5']
                print("\nValores do evento 5 após a conversão:")
                print(evento5_apos_conversao[[cod_evento_col, valor_col, 'NOME']].to_string())
            else:
                print("Nenhum registro encontrado para o evento 5")
        
        print("\nLendo arquivo Depara...")
        df_depara = read_file(depara_file)
        print(f"Total de registros no Depara: {len(df_depara)}")
        
        # Define as colunas específicas para o join
        lg_cod_col = "CÓDIGO EVENTO"
        depara_origem_col = "CÓDIGO EVENTO"
        depara_esocial_col = "COD ESOCIAL"
        
        # Verifica se as colunas existem nos arquivos
        colunas_lg = set(df_lg.columns)
        colunas_depara = set(df_depara.columns)
        
        # Encontra as colunas corretas
        lg_cod_col_atual = encontrar_coluna(lg_cod_col, colunas_lg)
        depara_origem_col_atual = encontrar_coluna(depara_origem_col, colunas_depara)
        depara_esocial_col_atual = encontrar_coluna(depara_esocial_col, colunas_depara)

        if not lg_cod_col_atual:
            raise ValueError(f"Coluna '{lg_cod_col}' não encontrada no arquivo LG")
        if not depara_origem_col_atual:
            raise ValueError(f"Coluna '{depara_origem_col}' não encontrada no arquivo Depara")
        if not depara_esocial_col_atual:
            raise ValueError(f"Coluna '{depara_esocial_col}' não encontrada no arquivo Depara")
            
        # Garante que as colunas de junção tenham o mesmo tipo de dados
        print("\nConvertendo tipos de dados para garantir compatibilidade...")
        df_lg[lg_cod_col_atual] = df_lg[lg_cod_col_atual].astype(str).str.strip()
        df_depara[depara_origem_col_atual] = df_depara[depara_origem_col_atual].astype(str).str.strip()
        
        # Exibe os tipos de dados das colunas de junção
        print(f"Tipo de dados da coluna '{lg_cod_col_atual}' no LG: {df_lg[lg_cod_col_atual].dtype}")
        print(f"Tipo de dados da coluna '{depara_origem_col_atual}' no Depara: {df_depara[depara_origem_col_atual].dtype}")
        
        # Exibe alguns valores de exemplo para verificação
        print("\nValores de exemplo da coluna de junção no LG:")
        print(df_lg[lg_cod_col_atual].head().to_string())
        print("\nValores de exemplo da coluna de junção no Depara:")
        print(df_depara[depara_origem_col_atual].head().to_string())
        
        # Realiza o merge entre LG e Depara
        print("\nRealizando merge entre LG e Depara...")
        df_merged = pd.merge(
            df_lg,
            df_depara,
            left_on=lg_cod_col_atual,
            right_on=depara_origem_col_atual,
            how='left'
        )
        
        # Agrupa por código do eSocial e soma os valores
        df_grouped = df_merged.groupby([depara_esocial_col_atual], as_index=False).agg({
            valor_col: 'sum'
        })
        
        # Ordena por código do eSocial
        df_grouped = df_grouped.sort_values(by=depara_esocial_col_atual)
        
        # Formata o valor para duas casas decimais
        df_grouped[valor_col] = df_grouped[valor_col].round(2)
        
        # Gera o nome do arquivo de saída com timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"lg_depara_agrupado_{timestamp}.xlsx"
        
        # Salva o resultado
        df_grouped.to_excel(output_file, index=False)
        print(f"\nArquivo salvo com sucesso: {output_file}")
        
        return df_grouped
        
    except Exception as e:
        print(f"\nErro ao processar arquivos: {str(e)}")
        return None

if __name__ == "__main__":
    # Arquivos de entrada
    lg_file = "Ficha Financeira.csv"
    depara_file = "DEPARA.xlsx"
    
    # Processa os arquivos
    result = process_lg_depara(lg_file, depara_file)
    
    if result is not None:
        print("\nProcessamento concluído com sucesso!") 