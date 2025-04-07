import pandas as pd
from datetime import datetime
import os
import re
import sys
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

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
                print( f"Falha na primeira tentativa: {str(e)}" )
                
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
            
            # Verifica se o valor está no formato brasileiro (com ponto como separador de milhar)
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
                # Verifica se o valor tem mais de 2 dígitos e não tem separador decimal
                if len(x) > 2:
                    # Assume que os últimos 2 dígitos são centavos
                    return float(x[:-2] + '.' + x[-2:])
                else:
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
                # Tenta diferentes formatos
                if '.' in valor_original and ',' in valor_original:
                    # Formato brasileiro: 1.234,56
                    valor_convertido = float(valor_original.replace('.', '').replace(',', '.'))
                elif ',' in valor_original:
                    # Apenas vírgula: 1234,56
                    valor_convertido = float(valor_original.replace(',', '.'))
                elif '.' in valor_original:
                    # Apenas ponto: 1234.56
                    valor_convertido = float(valor_original)
                else:
                    # Apenas números: 1234
                    # Verifica se o valor tem mais de 2 dígitos e não tem separador decimal
                    if len(valor_original) > 2:
                        # Assume que os últimos 2 dígitos são centavos
                        valor_convertido = float(valor_original[:-2] + '.' + valor_original[-2:])
                    else:
                        valor_convertido = float(valor_original)
                
                df.at[idx, valor_col] = valor_convertido
                print(f"Valor corrigido: {valor_original} -> {valor_convertido}")
            except:
                print(f"Não foi possível converter o valor: {valor_original}")
    
    # Preenche os valores NaN restantes com 0
    df[valor_col] = df[valor_col].fillna(0)
    
    return df

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
                    # Converte o valor para string e trata como valor monetário
                    valor_str = str(valor).strip()
                    if valor_str and valor_str.lower() != 'nan':
                        # Verifica se o valor é numérico
                        if isinstance(valor, (int, float)) or valor_str.replace(',', '').replace('.', '').isdigit():
                            # Se for numérico, converte para float
                            if isinstance(valor, (int, float)):
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
                                    # Verifica se o valor tem mais de 2 dígitos e não tem separador decimal
                                    if len(valor_str) > 2:
                                        # Assume que os últimos 2 dígitos são centavos
                                        valor_numerico = float(valor_str[:-2] + '.' + valor_str[-2:])
                                    else:
                                        valor_numerico = float(valor_str)
                            
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

def process_esocial_file(input_file):
    """
    Processa o arquivo eSocial e retorna os dados agrupados
    """
    try:
        print(f"Processando arquivo: {input_file}")
        
        # Lê o arquivo Excel
        df_esocial = pd.read_excel(input_file)
        print("Arquivo lido com sucesso")
        
        # Verifica se há colunas de valor e converte para o formato correto
        valor_cols = [col for col in df_esocial.columns if 'VALOR' in col.upper()]
        for col in valor_cols:
            print(f"Convertendo coluna de valor: {col}")
            df_esocial = convert_valor_to_numeric(df_esocial, col)
        
        # Transforma os dados
        df_transformed = transform_data(df_esocial)
        if df_transformed is None:
            raise ValueError("Erro na transformação dos dados")
        
        print("Dados transformados com sucesso")
        
        # Agrupa os dados
        df_grouped = group_and_sum(df_transformed)
        
        # Gera o nome do arquivo de saída com timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        """
        output_file = f"esocial_agrupado_{timestamp}.xlsx"
        
        # Salva o resultado
        df_grouped.to_excel(output_file, index=False)
        print(f"\nArquivo salvo com sucesso: {output_file}")
        
        """
        
        return df_grouped
        
    except Exception as e:
        print(f"\nErro ao processar arquivo: {str(e)}")
        return None

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
        
        # Define as colunas para a segunda planilha
        colunas_selecionadas = [
            'CÓDIGO EMPRESA', 'EMPRESA', 'REFERÊNCIA', 'CÓDIGO FOLHA',
            'FOLHA', 'CÓDIGO EVENTO', 'EVENTO', 'MATRÍCULA', 'NOME',
            'TIPO EVENTO', 'VALOR', depara_esocial_col_atual
        ]
        
        # Verifica se todas as colunas existem no DataFrame
        colunas_existentes = []
        for coluna in colunas_selecionadas:
            coluna_encontrada = encontrar_coluna(coluna, df_merged.columns)
            if coluna_encontrada:
                colunas_existentes.append(coluna_encontrada)
            else:
                print(f"Aviso: Coluna '{coluna}' não encontrada no DataFrame")
        
        # Cria o DataFrame com as colunas selecionadas
        df_selecionado = df_merged[colunas_existentes].copy()
        
        # Normaliza a coluna de tipo de evento (usado apenas para cálculos)
        tipo_evento_col = encontrar_coluna('TIPO EVENTO', df_selecionado.columns)
        df_selecionado['TIPO_CALC'] = df_selecionado[tipo_evento_col].str.upper().str.strip()
        
        # Cria uma coluna de valor ajustado (usado apenas para cálculos)
        df_selecionado['VALOR_AJUSTADO'] = df_selecionado.apply(
            lambda row: row[valor_col] if 'PROVENTO' in row['TIPO_CALC'] else -row[valor_col],
            axis=1
        )
        
        # Separa os registros com e sem código eSocial
        df_com_esocial = df_selecionado[df_selecionado[depara_esocial_col_atual].notna()].copy()
        df_sem_esocial = df_selecionado[df_selecionado[depara_esocial_col_atual].isna()].copy()
        
        # Define as colunas de agrupamento (sem CÓDIGO EVENTO e EVENTO)
        colunas_agrupamento = [
            'CÓDIGO EMPRESA', 'EMPRESA', 'CÓDIGO FOLHA', 'FOLHA',
            'MATRÍCULA', 'NOME', depara_esocial_col_atual
        ]
        
        # Verifica e ajusta os nomes das colunas de agrupamento
        colunas_agrupamento_atual = []
        for coluna in colunas_agrupamento:
            coluna_encontrada = encontrar_coluna(coluna, df_selecionado.columns)
            if coluna_encontrada:
                colunas_agrupamento_atual.append(coluna_encontrada)
            else:
                print(f"Aviso: Coluna '{coluna}' não encontrada para agrupamento")
        
        # Processa os registros com código eSocial (agrupados)
        print("\nAgrupando dados com código eSocial...")
        df_agrupado = df_com_esocial.groupby(colunas_agrupamento_atual).agg({
            'VALOR_AJUSTADO': 'sum',
            'TIPO_CALC': lambda x: ', '.join(sorted(set(x)))
        }).reset_index()
        
        # Adiciona as colunas CÓDIGO EVENTO e EVENTO vazias para registros agrupados
        df_agrupado[cod_evento_col] = '0'
        df_agrupado['EVENTO'] = ''
        
        # Renomeia as colunas
        df_agrupado = df_agrupado.rename(columns={
            'VALOR_AJUSTADO': 'VALOR_LIQUIDO',
            'TIPO_CALC': 'TIPOS_EVENTO'
        })
        
        # Calcula proventos e descontos separadamente para registros com eSocial
        proventos = df_com_esocial[df_com_esocial['TIPO_CALC'].str.contains('PROVENTO', na=False)]
        descontos = df_com_esocial[df_com_esocial['TIPO_CALC'].str.contains('DESCONTO', na=False)]
        
        # Soma proventos e descontos por grupo
        total_proventos = proventos.groupby(colunas_agrupamento_atual)[valor_col].sum()
        total_descontos = descontos.groupby(colunas_agrupamento_atual)[valor_col].sum()
        
        # Adiciona as colunas de totais ao DataFrame agrupado
        df_agrupado['TOTAL_PROVENTOS'] = df_agrupado.set_index(colunas_agrupamento_atual).index.map(total_proventos).fillna(0)
        df_agrupado['TOTAL_DESCONTOS'] = df_agrupado.set_index(colunas_agrupamento_atual).index.map(total_descontos).fillna(0)
        
        # Ajusta os descontos para serem positivos quando não houver proventos
        df_agrupado['TOTAL_DESCONTOS'] = df_agrupado.apply(
            lambda row: abs(row['TOTAL_DESCONTOS']) if row['TOTAL_PROVENTOS'] == 0 else row['TOTAL_DESCONTOS'],
            axis=1
        )
        
        # Calcula o VALOR_LIQUIDO para registros agrupados
        df_agrupado['VALOR_LIQUIDO'] = df_agrupado.apply(
            lambda row: (
                abs(row['TOTAL_DESCONTOS']) if row['TOTAL_PROVENTOS'] == 0
                else abs(row['TOTAL_PROVENTOS'] - abs(row['TOTAL_DESCONTOS']))
            ),
            axis=1
        )
        
        # Prepara os registros sem código eSocial (mantendo código evento e descrição originais)
        print("\nPreparando registros sem código eSocial...")
        df_sem_esocial['TIPOS_EVENTO'] = df_sem_esocial['TIPO_CALC']
        df_sem_esocial['TOTAL_PROVENTOS'] = df_sem_esocial.apply(
            lambda row: row[valor_col] if 'PROVENTO' in row['TIPO_CALC'] else 0,
            axis=1
        )
        df_sem_esocial['TOTAL_DESCONTOS'] = df_sem_esocial.apply(
            lambda row: abs(row[valor_col]) if 'DESCONTO' in row['TIPO_CALC'] else 0,
            axis=1
        )
        
        # Calcula o VALOR_LIQUIDO para registros não agrupados
        df_sem_esocial['VALOR_LIQUIDO'] = df_sem_esocial.apply(
            lambda row: (
                # Para eventos do tipo RESULTADO, mantém o valor original (sempre positivo)
                abs(row[valor_col]) if 'RESULTADO' in row['TIPO_CALC']
                # Para outros eventos, aplica a lógica de provento/desconto
                else (
                    abs(row['TOTAL_DESCONTOS']) if row['TOTAL_PROVENTOS'] == 0
                    else abs(row['TOTAL_PROVENTOS'] - abs(row['TOTAL_DESCONTOS']))
                )
            ),
            axis=1
        )
        
        # Reordena as colunas na ordem desejada
        colunas_finais = [
            'CÓDIGO EMPRESA', 'EMPRESA', 'CÓDIGO FOLHA', 'FOLHA',
            cod_evento_col, 'EVENTO', 'MATRÍCULA', 'NOME',
            depara_esocial_col_atual, 'TIPOS_EVENTO',
            'TOTAL_PROVENTOS', 'TOTAL_DESCONTOS', 'VALOR_LIQUIDO'
        ]
        
        # Garante que todas as colunas existam antes de reordenar
        for df in [df_sem_esocial, df_agrupado]:
            for col in colunas_finais:
                if col not in df.columns:
                    df[col] = ''  # Adiciona coluna vazia se não existir
        
        # Reordena as colunas
        df_sem_esocial = df_sem_esocial[colunas_finais]
        df_agrupado = df_agrupado[colunas_finais]
        
        # Combina os DataFrames (agrupados com eSocial + não agrupados sem eSocial)
        df_final = pd.concat([df_agrupado, df_sem_esocial], ignore_index=True)
        
        # Ordena o DataFrame final
        df_final = df_final.sort_values(by=['MATRÍCULA', depara_esocial_col_atual])
        
        # Agrupa por código do eSocial e soma os valores para a quarta planilha
        df_grouped = df_merged.groupby([depara_esocial_col_atual], as_index=False).agg({
            valor_col: 'sum'
        })
        
        # Ordena por código do eSocial
        df_grouped = df_grouped.sort_values(by=depara_esocial_col_atual)
        
        # Formata o valor para duas casas decimais
        df_grouped[valor_col] = df_grouped[valor_col].round(2)
        
        """
        # Gera o nome do arquivo de saída com timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"lg_depara_agrupado_{timestamp}.xlsx"
        
        # Salva o resultado
        df_grouped.to_excel(output_file, index=False)
        print(f"Arquivo salvo com sucesso: {output_file}")
        """
        
        # Retorna um dicionário com todos os DataFrames
        return {
            'dados_completos': df_merged[colunas_existentes],
            'dados_selecionados': df_selecionado.drop(columns=['TIPO_CALC', 'VALOR_AJUSTADO']),
            'valores_por_esocial': df_final,
            'agrupado_por_esocial': df_grouped
        }
        
    except Exception as e:
        print(f"\nErro ao processar arquivos: {str(e)}")
        return None

def main():
    try:
        # Verifica se os argumentos foram fornecidos
        if len(sys.argv) < 4:
            print("Uso: python processador_completo.py <arquivo_esocial> <arquivo_ficha_financeira> <arquivo_depara>")
            print("Exemplo: python processador_completo.py ANDREANI_dirf_x_esocial_1.xlsx ficha_financeira.csv depara.csv")
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
        result_lg_depara = process_lg_depara(lg_file, depara_file)
        if result_lg_depara is None:
            raise ValueError("Erro ao processar arquivos LG e Depara")
        
        # Prepara os dados para o join entre eSocial e Valores por eSocial
        print("\n3. Preparando dados para comparação...")
        
        # Exibe as colunas disponíveis
        print("\nColunas disponíveis no eSocial:", df_esocial.columns.tolist())
        print("\nColunas disponíveis em Valores por eSocial:", result_lg_depara['valores_por_esocial'].columns.tolist())
        
        # Prepara o DataFrame do eSocial
        df_esocial_join = df_esocial.copy()
        
        # Converte as colunas para string e faz trim
        df_esocial_join['MATRICULA'] = df_esocial_join['MATRICULA'].astype(str).str.strip()
        df_esocial_join['Código'] = df_esocial_join['Código'].astype(str).str.strip()
        
        # Agrupa registros duplicados no eSocial somando os valores
        print("\nAgrupando registros duplicados no eSocial...")
        df_esocial_join = df_esocial_join.groupby(
            ['MATRICULA', 'Código', 'Descrição Completa'],
            as_index=False
        ).agg({
            'Valor': 'sum'
        })
        
        # Prepara o DataFrame de Valores por eSocial
        df_valores_esocial = result_lg_depara['valores_por_esocial'].copy()
        df_valores_esocial['MATRÍCULA'] = df_valores_esocial['MATRÍCULA'].astype(str).str.strip()
        df_valores_esocial['COD ESOCIAL'] = df_valores_esocial['COD ESOCIAL'].astype(str).str.strip()
        
        # Filtra apenas registros com COD ESOCIAL preenchido
        df_valores_esocial_com_cod = df_valores_esocial[df_valores_esocial['COD ESOCIAL'].notna() & (df_valores_esocial['COD ESOCIAL'] != '')]
        df_valores_esocial_sem_cod = df_valores_esocial[df_valores_esocial['COD ESOCIAL'].isna() | (df_valores_esocial['COD ESOCIAL'] == '')]
        
        # Exibe alguns valores antes do merge para diagnóstico
        print("\nExemplo de valores no eSocial após agrupamento:")
        print(df_esocial_join[['MATRICULA', 'Código', 'Descrição Completa', 'Valor']].head())
        print("\nExemplo de valores em Valores por eSocial (com código):")
        print(df_valores_esocial_com_cod[['MATRÍCULA', 'COD ESOCIAL']].head())
        
        # Verifica se ainda existem duplicatas após o agrupamento
        duplicatas = df_esocial_join.groupby(['MATRICULA', 'Código']).size().reset_index(name='count')
        duplicatas = duplicatas[duplicatas['count'] > 1]
        if not duplicatas.empty:
            print("\nAVISO: Ainda existem duplicatas após o agrupamento:")
            print(duplicatas)
            
            # Se ainda houver duplicatas, agrupa novamente sem considerar a Descrição Completa
            print("\nAgrupando novamente sem considerar a Descrição Completa...")
            df_esocial_join = df_esocial_join.groupby(
                ['MATRICULA', 'Código'],
                as_index=False
            ).agg({
                'Valor': 'sum',
                'Descrição Completa': 'first'  # Mantém a primeira descrição encontrada
            })
        
        # Realiza o merge apenas para registros com código eSocial
        print("\nRealizando merge entre eSocial e Valores por eSocial...")
        df_merged = pd.merge(
            df_valores_esocial_com_cod,
            df_esocial_join[['MATRICULA', 'Código', 'Descrição Completa', 'Valor']],
            left_on=['MATRÍCULA', 'COD ESOCIAL'],
            right_on=['MATRICULA', 'Código'],
            how='left',
            validate='many_to_one'
        )
        
        # Remove colunas duplicadas do join
        if 'MATRICULA' in df_merged.columns:
            df_merged = df_merged.drop(columns=['MATRICULA'])
        if 'Código' in df_merged.columns:
            df_merged = df_merged.drop(columns=['Código'])
        
        # Renomeia as colunas do eSocial
        df_merged = df_merged.rename(columns={
            'Valor': 'Valor_eSocial',
            'Descrição Completa': 'Descrição_Completa_eSocial'
        })
        
        # Converte os valores monetários para o formato correto
        if 'Valor_eSocial' in df_merged.columns:
            df_merged['Valor_eSocial'] = pd.to_numeric(df_merged['Valor_eSocial'], errors='coerce')
            df_merged['Valor_eSocial'] = df_merged['Valor_eSocial'].fillna(0).round(2)
        
        # Concatena os resultados (registros com e sem código eSocial)
        df_final = pd.concat([df_valores_esocial_sem_cod, df_merged], ignore_index=True)
        
        # Ordena o DataFrame final por MATRÍCULA e COD ESOCIAL
        df_final = df_final.sort_values(['MATRÍCULA', 'COD ESOCIAL'])
        
        # Atualiza o DataFrame de Valores por eSocial
        result_lg_depara['valores_por_esocial'] = df_final
        
        # Gera o nome do arquivo de saída com timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"resultado_comparacao_{timestamp}.xlsx"
        
        # Salva os resultados em diferentes abas do mesmo arquivo Excel
        print(f"\nSalvando resultados em {output_file}...")
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # Salva as planilhas do LG/Depara
            result_lg_depara['dados_completos'].to_excel(writer, sheet_name='Dados Completos', index=False)
            result_lg_depara['dados_selecionados'].to_excel(writer, sheet_name='Dados Selecionados', index=False)
            result_lg_depara['valores_por_esocial'].to_excel(writer, sheet_name='Valores por eSocial', index=False)
            
            # Salva a planilha do eSocial
            df_esocial.to_excel(writer, sheet_name='eSocial', index=False)
            
            # Formata a planilha "Dados Selecionados"
            worksheet = writer.sheets['Dados Selecionados']
            tipo_evento_col = None
            for idx, col in enumerate(result_lg_depara['dados_selecionados'].columns):
                if 'TIPO EVENTO' in col.upper():
                    tipo_evento_col = idx
                    break
            
            if tipo_evento_col is not None:
                cor_provento = PatternFill(start_color='90EE90', end_color='90EE90', fill_type='solid')  # Verde claro
                cor_desconto = PatternFill(start_color='FFB6C1', end_color='FFB6C1', fill_type='solid')  # Vermelho claro
                
                for row_idx, row in enumerate(worksheet.iter_rows(min_row=2), start=2):
                    tipo_evento = str(row[tipo_evento_col].value).upper() if row[tipo_evento_col].value else ""
                    cor = None
                    if "PROVENTO" in tipo_evento:
                        cor = cor_provento
                    elif "DESCONTO" in tipo_evento:
                        cor = cor_desconto
                    if cor:
                        for cell in row:
                            cell.fill = cor
            
            # Ajusta a largura das colunas em todas as planilhas
            for sheet_name in writer.sheets:
                worksheet = writer.sheets[sheet_name]
                if sheet_name == 'Dados Completos':
                    df_to_use = result_lg_depara['dados_completos']
                elif sheet_name == 'Dados Selecionados':
                    df_to_use = result_lg_depara['dados_selecionados']
                elif sheet_name == 'Valores por eSocial':
                    df_to_use = result_lg_depara['valores_por_esocial']
                else:  # eSocial
                    df_to_use = df_esocial
                
                for idx, col in enumerate(df_to_use.columns):
                    max_length = max(
                        df_to_use[col].astype(str).apply(len).max(),
                        len(str(col))
                    )
                    worksheet.column_dimensions[get_column_letter(idx + 1)].width = max_length + 2
                    
                    # Formata as colunas de valores como número com duas casas decimais
                    if col in ['VALOR', 'TOTAL_PROVENTOS', 'TOTAL_DESCONTOS', 'VALOR_LIQUIDO', 'Valor_eSocial']:
                        for cell in worksheet[get_column_letter(idx + 1)][1:]:
                            cell.number_format = '#,##0.00'
                            
                  
        
        print("\nProcessamento concluído com sucesso!")
        print(f"Arquivo de saída: {output_file}")
        
    except Exception as e:
        print(f"\nErro durante o processamento: {str(e)}")

if __name__ == "__main__":
    main() 