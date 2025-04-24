import pandas as pd
import sqlite3
import os
from datetime import datetime
import sys
import unicodedata

def normalizar_nome_coluna(nome):
    """
    Normaliza o nome da coluna removendo acentos e caracteres especiais
    """
    # Remove acentos
    nome = unicodedata.normalize('NFKD', nome).encode('ASCII', 'ignore').decode('ASCII')
    # Converte para maiúsculas e remove espaços extras
    nome = nome.upper().strip()
    return nome

def normalizar_colunas_df(df):
    """
    Normaliza os nomes de todas as colunas do DataFrame
    """
    # Cria um dicionário com o mapeamento dos nomes originais para os normalizados
    mapeamento = {col: normalizar_nome_coluna(col) for col in df.columns}
    # Renomeia as colunas
    return df.rename(columns=mapeamento)

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
                    print(f"Arquivo {file_path} lido com sucesso usando cp1252")
                    return df
            except Exception as e:
                print(f"Falha na primeira tentativa: {str(e)}")
                
                # Segunda tentativa: tentar outros encodings comuns
                encodings = ['iso-8859-1', 'latin1', 'utf-8-sig', 'utf-8', 'utf-16']
                
                for encoding in encodings:
                    try:
                        print(f"Tentando ler com encoding: {encoding}")
                        df = pd.read_csv(file_path, sep=';', encoding=encoding, low_memory=False)
                        print(f"Arquivo {file_path} lido com sucesso usando encoding: {encoding}")
                        return df
                    except Exception as e:
                        print(f"Falha com encoding {encoding}: {str(e)}")
                        continue
                
                raise ValueError(f"Não foi possível ler o arquivo {file_path} com nenhum dos encodings disponíveis")
            
        elif file_extension in ['.xlsx', '.xls']:
            df = pd.read_excel(file_path)
            print(f"Arquivo {file_path} lido com sucesso")
            return df
        else:
            raise ValueError(f"Formato de arquivo não suportado: {file_extension}")
    except Exception as e:
        raise Exception(f"Erro ao ler o arquivo {file_path}: {str(e)}")

def limpar_tabelas(conn):
    """
    Limpa todos os registros das tabelas existentes
    """
    try:
        cursor = conn.cursor()
        
        # Obtém lista de todas as tabelas
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        tabelas = cursor.fetchall()
        
        # Limpa cada tabela
        for tabela in tabelas:
            nome_tabela = tabela[0]
            if nome_tabela != 'sqlite_sequence':  # Ignora tabela de sistema
                print(f"Limpando tabela {nome_tabela}...")
                cursor.execute(f"DELETE FROM {nome_tabela}")
                # Reseta o contador de ID se existir
                cursor.execute(f"DELETE FROM sqlite_sequence WHERE name='{nome_tabela}'")
        
        conn.commit()
        print("Todas as tabelas foram limpas com sucesso!")
        
    except Exception as e:
        print(f"Erro ao limpar tabelas: {str(e)}")
        raise e

def create_database():
    """
    Cria o banco de dados SQLite e as tabelas necessárias
    """
    try:
        # Conecta ao banco de dados (cria se não existir)
        conn = sqlite3.connect('dados_esocial.db')
        cursor = conn.cursor()
        
        # Verifica se as tabelas já existem
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        tabelas_existentes = cursor.fetchall()
        
        # Se existirem tabelas, limpa os registros
        if tabelas_existentes:
            print("Tabelas existentes encontradas. Limpando registros...")
            limpar_tabelas(conn)
        else:
            print("Criando novas tabelas...")
            # Cria a tabela depara_eventos
            cursor.execute('''
            CREATE TABLE IF NOT EXISTS depara_eventos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                codigo TEXT,
                descricao TEXT,
                depara_esocial TEXT,
                data_importacao DATETIME DEFAULT CURRENT_TIMESTAMP
            )
            ''')
            
            # Cria a tabela depara
            cursor.execute('''
            CREATE TABLE IF NOT EXISTS depara (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                codigo_evento TEXT,
                cod_esocial TEXT,
                data_importacao DATETIME DEFAULT CURRENT_TIMESTAMP
            )
            ''')
            
            # Cria a tabela ficha_financeira
            cursor.execute('''
            CREATE TABLE IF NOT EXISTS ficha_financeira (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                codigo_empresa TEXT,
                empresa TEXT,
                referencia TEXT,
                codigo_folha TEXT,
                folha TEXT,
                codigo_evento TEXT,
                evento TEXT,
                matricula TEXT,
                nome TEXT,
                tipo_evento TEXT,
                valor REAL,
                data_pagamento DATE,
                data_importacao DATETIME DEFAULT CURRENT_TIMESTAMP
            )
            ''')
            
            # Cria a tabela esocial
            cursor.execute('''
            CREATE TABLE IF NOT EXISTS esocial (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                codigo_empresa TEXT,
                empresa TEXT,
                referencia TEXT,
                codigo_folha TEXT,
                folha TEXT,
                matricula TEXT,
                matricula_esocial TEXT,
                cpf TEXT,
                nome TEXT,
                data_admissao DATE,
                codigo_tarefa TEXT,
                nome_tarefa TEXT,
                data_processamento DATE,
                hora_processamento TEXT,
                id_execucao TEXT,
                tipo_acao TEXT,
                status_registro TEXT,
                ambiente_esocial TEXT,
                numero_recibo TEXT,
                indicador_periodo TEXT,
                periodo_apuracao TEXT,
                periodo_referencia TEXT,
                demonstrativo_valores TEXT,
                tipo_pagamento TEXT,
                data_pagamento DATE,
                categoria_trabalhador TEXT,
                codigo TEXT,
                descricao TEXT,
                descricao_completa TEXT,
                valor REAL,
                data_importacao DATETIME DEFAULT CURRENT_TIMESTAMP
            )
            ''')
            
            conn.commit()
            print("Banco de dados e tabelas criados com sucesso!")
        
        return conn
        
    except Exception as e:
        print(f"Erro ao criar banco de dados: {str(e)}")
        raise e

def import_depara_eventos(conn, file_path):
    """
    Importa dados do arquivo depara_eventos_esocial.csv
    """
    try:
        df = read_file(file_path)
        
        # Exibe informações sobre o DataFrame
        print(f"\nColunas no arquivo {file_path}:")
        print(df.columns.tolist())
        print(f"Total de registros: {len(df)}")
        
        # Normaliza os nomes das colunas
        df = normalizar_colunas_df(df)
        print("\nColunas após normalização:")
        print(df.columns.tolist())
        
        # Mapeia as colunas para os nomes esperados
        column_mapping = {}
        for col in df.columns:
            if 'CODIGO' in col:
                column_mapping[col] = 'codigo'
            elif 'DESCRICAO' in col:
                column_mapping[col] = 'descricao'
            elif 'DEPARA' in col or 'ESOCIAL' in col:
                column_mapping[col] = 'depara_esocial'
        
        # Renomeia as colunas
        df = df.rename(columns=column_mapping)
        
        # Adiciona coluna de data de importação
        df['data_importacao'] = datetime.now()
        
        # Importa para o banco de dados
        df.to_sql('depara_eventos', conn, if_exists='append', index=False)
        print(f"Dados do arquivo {file_path} importados com sucesso")
        
    except Exception as e:
        print(f"Erro ao importar dados do arquivo {file_path}: {str(e)}")

def import_depara(conn, file_path):
    """
    Importa dados do arquivo DEPARA.xlsx
    """
    try:
        df = read_file(file_path)
        
        # Exibe informações sobre o DataFrame
        print(f"\nColunas no arquivo {file_path}:")
        print(df.columns.tolist())
        print(f"Total de registros: {len(df)}")
        
        # Normaliza os nomes das colunas
        df = normalizar_colunas_df(df)
        print("\nColunas após normalização:")
        print(df.columns.tolist())
        
        # Mapeia as colunas para os nomes esperados
        column_mapping = {}
        for col in df.columns:
            if 'CODIGO' in col and 'EVENTO' in col:
                column_mapping[col] = 'codigo_evento'
            elif 'COD' in col and 'ESOCIAL' in col:
                column_mapping[col] = 'cod_esocial'
        
        # Renomeia as colunas
        df = df.rename(columns=column_mapping)
        
        # Adiciona coluna de data de importação
        df['data_importacao'] = datetime.now()
        
        # Importa para o banco de dados
        df.to_sql('depara', conn, if_exists='append', index=False)
        print(f"Dados do arquivo {file_path} importados com sucesso")
        
    except Exception as e:
        print(f"Erro ao importar dados do arquivo {file_path}: {str(e)}")

def import_ficha_financeira(conn, file_path):
    """
    Importa dados do arquivo Ficha Financeira.csv
    """
    try:
        df = read_file(file_path)
        
        # Exibe informações sobre o DataFrame
        print(f"\nColunas no arquivo {file_path}:")
        print(df.columns.tolist())
        print(f"Total de registros: {len(df)}")
        
        # Normaliza os nomes das colunas
        df = normalizar_colunas_df(df)
        print("\nColunas após normalização:")
        print(df.columns.tolist())
        
        # Mapeia as colunas para os nomes esperados
        column_mapping = {}
        for col in df.columns:
            if 'CODIGO' in col and 'EMPRESA' in col:
                column_mapping[col] = 'codigo_empresa'
            elif 'EMPRESA' in col and 'CODIGO' not in col:
                column_mapping[col] = 'empresa'
            elif 'REFERENCIA' in col:
                column_mapping[col] = 'referencia'
            elif 'CODIGO' in col and 'FOLHA' in col:
                column_mapping[col] = 'codigo_folha'
            elif 'FOLHA' in col and 'CODIGO' not in col:
                column_mapping[col] = 'folha'
            elif 'CODIGO' in col and 'EVENTO' in col:
                column_mapping[col] = 'codigo_evento'
            elif 'EVENTO' in col and 'CODIGO' not in col and 'TIPO' not in col:
                column_mapping[col] = 'evento'
            elif 'MATRICULA' in col:
                column_mapping[col] = 'matricula'
            elif 'NOME' in col:
                column_mapping[col] = 'nome'
            elif 'TIPO' in col and 'EVENTO' in col:
                column_mapping[col] = 'tipo_evento'
            elif 'VALOR' in col:
                column_mapping[col] = 'valor'
            elif 'DATA' in col and 'PAGAMENTO' in col:
                column_mapping[col] = 'data_pagamento'
        
        # Renomeia as colunas
        df = df.rename(columns=column_mapping)
        
        # Converte a coluna valor para numérico
        if 'valor' in df.columns:
            # Função para converter valores no formato brasileiro
            def convert_br_number(x):
                try:
                    if pd.isna(x) or x == '':
                        return 0.0
                    # Converte para string e remove espaços
                    x = str(x).strip()
                    # Remove pontos de milhar e substitui vírgula por ponto
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
                    return 0.0
            
            # Aplica a conversão
            df['valor'] = df['valor'].apply(convert_br_number)
            
            # Exibe alguns valores para verificação
            print("\nExemplos de valores convertidos:")
            print(df[['matricula', 'nome', 'valor']].head())
        
        # Adiciona coluna de data de importação
        df['data_importacao'] = datetime.now()
        
        # Importa para o banco de dados
        df.to_sql('ficha_financeira', conn, if_exists='append', index=False)
        print(f"Dados do arquivo {file_path} importados com sucesso")
        
    except Exception as e:
        print(f"Erro ao importar dados do arquivo {file_path}: {str(e)}")

def import_esocial(conn, file_path):
    """
    Importa dados do arquivo esocial.xlsx com transformação dos dados
    """
    try:
        df = read_file(file_path)
        
        # Exibe informações sobre o DataFrame
        print(f"\nColunas no arquivo {file_path}:")
        print(df.columns.tolist())
        print(f"Total de registros: {len(df)}")
        
        # Normaliza os nomes das colunas
        df = normalizar_colunas_df(df)
        print("\nColunas após normalização:")
        print(df.columns.tolist())
        
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
                                # Extrai código e descrição da coluna
                                partes = coluna.split(' - ', 1)
                                if len(partes) == 2:
                                    codigo = partes[0].strip()
                                    descricao = partes[1].strip()
                                else:
                                    codigo = ''
                                    descricao = coluna.strip()
                                
                                nova_linha['codigo'] = codigo
                                nova_linha['descricao'] = descricao
                                nova_linha['descricao_completa'] = coluna
                                nova_linha['valor'] = valor_numerico
                                linhas_transformadas.append(nova_linha)
                except (ValueError, TypeError):
                    continue
        
        if not linhas_transformadas:
            print("Nenhum dado foi transformado")
            return
        
        # Cria o DataFrame com as linhas transformadas
        df_transformed = pd.DataFrame(linhas_transformadas)
        
        # Mapeia as colunas fixas para os nomes esperados
        column_mapping = {}
        for col in df_transformed.columns:
            if 'CODIGO' in col and 'EMPRESA' in col:
                column_mapping[col] = 'codigo_empresa'
            elif 'EMPRESA' in col and 'CODIGO' not in col:
                column_mapping[col] = 'empresa'
            elif 'REFERENCIA' in col:
                column_mapping[col] = 'referencia'
            elif 'CODIGO' in col and 'FOLHA' in col:
                column_mapping[col] = 'codigo_folha'
            elif 'FOLHA' in col and 'CODIGO' not in col:
                column_mapping[col] = 'folha'
            elif 'MATRICULA' in col:
                column_mapping[col] = 'matricula'
            elif 'NOME' in col:
                column_mapping[col] = 'nome'
            elif 'CATEGORIA' in col and 'TRABALHADOR' in col:
                column_mapping[col] = 'categoria_trabalhador'
            elif 'CPF' in col:
                column_mapping[col] = 'cpf'
            elif 'DATA' in col and 'ADMISSAO' in col:
                column_mapping[col] = 'data_admissao'
            elif 'CODIGO' in col and 'TAREFA' in col:
                column_mapping[col] = 'codigo_tarefa'
            elif 'NOME' in col and 'TAREFA' in col:
                column_mapping[col] = 'nome_tarefa'
            elif 'DATA' in col and 'PROCESSAMENTO' in col:
                column_mapping[col] = 'data_processamento'
            elif 'HORA' in col and 'PROCESSAMENTO' in col:
                column_mapping[col] = 'hora_processamento'
            elif 'ID' in col and 'EXECUCAO' in col:
                column_mapping[col] = 'id_execucao'
            elif 'TIPO' in col and 'ACAO' in col:
                column_mapping[col] = 'tipo_acao'
            elif 'STATUS' in col and 'REGISTRO' in col:
                column_mapping[col] = 'status_registro'
            elif 'AMBIENTE' in col and 'ESOCIAL' in col:
                column_mapping[col] = 'ambiente_esocial'
            elif 'NUMERO' in col and 'RECIBO' in col:
                column_mapping[col] = 'numero_recibo'
            elif 'INDICADOR' in col and 'PERIODO' in col:
                column_mapping[col] = 'indicador_periodo'
            elif 'PERIODO' in col and 'APURACAO' in col:
                column_mapping[col] = 'periodo_apuracao'
            elif 'PERIODO' in col and 'REFERENCIA' in col:
                column_mapping[col] = 'periodo_referencia'
            elif 'DEMONSTRATIVO' in col and 'VALORES' in col:
                column_mapping[col] = 'demonstrativo_valores'
            elif 'TIPO' in col and 'PAGAMENTO' in col:
                column_mapping[col] = 'tipo_pagamento'
            elif 'DATA' in col and 'PAGAMENTO' in col:
                column_mapping[col] = 'data_pagamento'
        
        # Renomeia as colunas
        df_transformed = df_transformed.rename(columns=column_mapping)
        
        # Adiciona coluna de data de importação
        df_transformed['data_importacao'] = datetime.now()
        
        # Importa para o banco de dados
        df_transformed.to_sql('esocial', conn, if_exists='append', index=False)
        print(f"Dados do arquivo {file_path} importados com sucesso")
        
    except Exception as e:
        print(f"Erro ao importar dados do arquivo {file_path}: {str(e)}")

def main():
    try:
        # Define os nomes dos arquivos diretamente
        depara_eventos_file = "depara_eventos_esocial.csv"
        depara_file = "DEPARA.xlsx"
        ficha_financeira_file = "Ficha Financeira.CSV"
        esocial_file = "esocial.xlsx"
        
        # Verifica se os arquivos existem
        for file_path in [depara_eventos_file, depara_file, ficha_financeira_file, esocial_file]:
            if not os.path.exists(file_path):
                print(f"Erro: Arquivo não encontrado: {file_path}")
                return
        
        print("=== Iniciando importação de dados ===")
        print(f"Arquivo Depara Eventos: {depara_eventos_file}")
        print(f"Arquivo Depara: {depara_file}")
        print(f"Arquivo Ficha Financeira: {ficha_financeira_file}")
        print(f"Arquivo eSocial: {esocial_file}")
        
        # Cria o banco de dados
        conn = create_database()
        if conn is None:
            return
        
        # Importa os dados
        import_depara_eventos(conn, depara_eventos_file)
        import_depara(conn, depara_file)
        import_ficha_financeira(conn, ficha_financeira_file)
        import_esocial(conn, esocial_file)
        
        # Fecha a conexão com o banco de dados
        conn.close()
        print("\nImportação concluída com sucesso!")
        
    except Exception as e:
        print(f"\nErro durante a importação: {str(e)}")

if __name__ == "__main__":
    main() 