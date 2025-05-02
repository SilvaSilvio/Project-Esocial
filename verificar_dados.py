import sqlite3
import pandas as pd
from datetime import datetime
import os
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
import sys

def print_with_timestamp(message):
    """Função para imprimir mensagens com timestamp"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] {message}")
    sys.stdout.flush()

def verificar_dados():
    """
    Verifica os dados importados no banco de dados
    """
    try:
        # Conecta ao banco de dados
        conn = sqlite3.connect('dados_esocial.db')
        cursor = conn.cursor()
        
        # Obtém a lista de tabelas
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        tabelas = cursor.fetchall()
        
        print_with_timestamp("=== Verificação de Dados ===")
        
        # Para cada tabela, mostra o número de registros e alguns exemplos
        for tabela in tabelas:
            nome_tabela = tabela[0]
            if nome_tabela != 'sqlite_sequence':  # Ignora tabela de sistema
                # Conta registros
                cursor.execute(f"SELECT COUNT(*) FROM {nome_tabela}")
                total_registros = cursor.fetchone()[0]
                
                print_with_timestamp(f"\nTabela: {nome_tabela}")
                print_with_timestamp(f"Total de registros: {total_registros}")
                
                # Mostra alguns registros como exemplo
                if total_registros > 0:
                    cursor.execute(f"SELECT * FROM {nome_tabela} LIMIT 5")
                    registros = cursor.fetchall()
                    
                    # Obtém os nomes das colunas
                    cursor.execute(f"PRAGMA table_info({nome_tabela})")
                    colunas = [col[1] for col in cursor.fetchall()]
                    
                    # Cria um DataFrame para exibição
                    df = pd.DataFrame(registros, columns=colunas)
                    print_with_timestamp("\nExemplos de registros:")
                    print(df)
        
        conn.close()
        print_with_timestamp("\nVerificação concluída!")
        
    except Exception as e:
        print_with_timestamp(f"Erro ao verificar dados: {str(e)}")

def consultar_dados_join():
    """
    Executa consultas SQL com JOIN entre as tabelas e gera planilhas com os resultados
    """
    try:
        # Conecta ao banco de dados
        conn = sqlite3.connect('dados_esocial.db')
        
        print_with_timestamp("\n=== Executando Consultas com JOIN ===")
        
        '''# Consulta 1: Comparação detalhada entre ficha_financeira, depara e esocial
        query1 = """
        SELECT 
            ff.codigo_empresa,
            ff.empresa,
            ff.matricula,
            ff.nome,
            ff.codigo_folha,
            ff.folha,
            ff.codigo_evento,
            ff.evento,
            ff.tipo_evento,
            dp.cod_esocial,
            dpeve.depara_esocial,
            ff.valor,
            CASE 
                WHEN ff.tipo_evento = 'Provento' THEN ff.valor 
                ELSE 0 
            END as valor_proventos,
            CASE 
                WHEN ff.tipo_evento = 'Desconto' THEN ff.valor 
                ELSE 0 
            END as valor_descontos
        FROM ficha_financeira ff
        LEFT JOIN depara dp ON ff.codigo_evento = dp.codigo_evento
        LEFT JOIN depara_eventos dpeve ON ff.codigo_evento = dpeve.codigo
        WHERE ff.valor != 0
        ORDER BY ff.matricula, ff.codigo_evento
        """
        '''
        
        # Consulta 2: Valor agrupado por Código Esocial com totais
        query2 = """
        SELECT 
            ff.codigo_empresa,
            ff.empresa,
            ff.matricula,
            ff.nome,
            ff.codigo_folha,
            ff.folha,
            dp.cod_esocial,
            SUM(CASE WHEN ff.tipo_evento = 'Provento' THEN ff.valor ELSE 0 END) as total_proventos,
            SUM(CASE WHEN ff.tipo_evento = 'Desconto' THEN ff.valor ELSE 0 END) as total_descontos,
            SUM(CASE WHEN ff.tipo_evento = 'Resultado' THEN ff.valor ELSE 0 END) as total_Resultado,
            ROUND(ABS(SUM(CASE WHEN ff.tipo_evento = 'Provento' THEN ff.valor ELSE 0 END) -
                 SUM(CASE WHEN ff.tipo_evento = 'Desconto' THEN ff.valor ELSE 0 END)), 2) as valor_liquido,
            (
            select
                    es.valor
            from esocial es
            where es.matricula = ff.matricula
              and replace( SUBSTR(es.descricao_completa, 1, INSTR(es.descricao_completa, ':') - 1),"TIPO-","") = dp.cod_esocial
            
            ) AS VALOR_ESOCIAL,

            (round( abs(( SUM(CASE WHEN ff.tipo_evento = 'Provento' THEN ff.valor ELSE 0 END)
            - SUM(CASE WHEN ff.tipo_evento = 'Desconto' THEN ff.valor ELSE 0 END)
            ) ), 2)
            -
            (
            select
                    es.valor
            from esocial es
            where es.matricula = ff.matricula
              and replace( SUBSTR(es.descricao_completa, 1, INSTR(es.descricao_completa, ':') - 1),"TIPO-","") = dp.cod_esocial
            
            ) ) AS Resultado_Final

        FROM ficha_financeira ff
        LEFT JOIN depara dp ON ff.codigo_evento = dp.codigo_evento
        LEFT JOIN depara_eventos dpeve ON ff.codigo_evento = dpeve.codigo
        WHERE ff.valor != 0
        GROUP BY
            ff.codigo_empresa,
            ff.empresa,
            ff.matricula,
            ff.nome,
            ff.codigo_folha,
            ff.folha,
            dp.cod_esocial
        ORDER BY ff.matricula, dp.cod_esocial
        """
   
        query3 = """
        SELECT 
            ff.codigo_empresa,
            ff.empresa,
            ff.matricula,
            ff.nome,
            ff.codigo_folha,
            ff.folha,
            dpeve.depara_esocial,
            SUM(CASE WHEN ff.tipo_evento = 'Provento' THEN ff.valor ELSE 0 END) as total_proventos,
            SUM(CASE WHEN ff.tipo_evento = 'Desconto' THEN ff.valor ELSE 0 END) as total_descontos,
            SUM(CASE WHEN ff.tipo_evento = 'Resultado' THEN ff.valor ELSE 0 END) as total_Resultado,
           round( abs(( SUM(CASE WHEN ff.tipo_evento = 'Provento' THEN ff.valor ELSE 0 END)
            - SUM(CASE WHEN ff.tipo_evento = 'Desconto' THEN ff.valor ELSE 0 END)
            ) ), 2)          
                as valor_liquido,
            round((SELECT sum( esocial.valor)
             FROM esocial esocial
             WHERE esocial.matricula = ff.matricula
             AND esocial.descricao_completa = dpeve.depara_esocial
             ), 2 ) as valor_esocial,
             
            (round( abs(( SUM(CASE WHEN ff.tipo_evento = 'Provento' THEN ff.valor ELSE 0 END)
            - SUM(CASE WHEN ff.tipo_evento = 'Desconto' THEN ff.valor ELSE 0 END)
            ) ), 2)          
            -
            round((SELECT sum( esocial.valor)
             FROM esocial esocial
             WHERE esocial.matricula = ff.matricula
             AND esocial.descricao_completa = dpeve.depara_esocial
             ), 2 ) ) as Resultado_Final
             
        FROM ficha_financeira ff
        LEFT JOIN depara dp ON ff.codigo_evento = dp.codigo_evento
        LEFT JOIN depara_eventos dpeve ON ff.codigo_evento = dpeve.codigo
        WHERE ff.valor != 0

        GROUP BY
            ff.codigo_empresa,
            ff.empresa,
            ff.matricula,
            ff.nome,
            ff.codigo_folha,
            ff.folha,
            dpeve.depara_esocial
        ORDER BY ff.matricula, dp.cod_esocial"""
        
        query4 = """
        select
            id,
            codigo_empresa,
            empresa,
            referencia,
            codigo_folha,
            folha,
            matricula,
            matricula_esocial,
            cpf,
            nome,
            data_admissao,
            codigo_tarefa,
            nome_tarefa,
            data_processamento,
            hora_processamento,
            id_execucao,
            tipo_acao,
            status_registro,
            ambiente_esocial,
            numero_recibo,
            indicador_periodo,
            periodo_apuracao,
            periodo_referencia,
            demonstrativo_valores,
            tipo_pagamento,
            data_pagamento,
            categoria_trabalhador,
            codigo,
            descricao,
            descricao_completa,
            valor,
            data_importacao
        from esocial 
        where descricao_completa not in ("CNPJ OPERADORA SAUDE", "CPF DEPENDENTE","PENSIONISTA - CPF", "IR COMPETENCIA - CPF DEPENDENTE")
        """
        
        # Executa as consultas e converte para DataFrames
        #print_with_timestamp("Executando consulta 1...")
        #df1 = pd.read_sql_query(query1, conn)
        #print_with_timestamp(f"Consulta 1 concluída. Total de registros: {len(df1)}")
        
        print_with_timestamp("Executando consulta 2...")
        df2 = pd.read_sql_query(query2, conn)
        print_with_timestamp(f"Consulta 2 concluída. Total de registros: {len(df2)}")
        
        print_with_timestamp("Executando consulta 3...")
        df3 = pd.read_sql_query(query3, conn)
        print_with_timestamp(f"Consulta 3 concluída. Total de registros: {len(df3)}")
        
        print_with_timestamp("Executando consulta 4...")
        df4 = pd.read_sql_query(query4, conn)
        print_with_timestamp(f"Consulta 4 concluída. Total de registros: {len(df4)}")
        
        # Gera nome do arquivo com timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"resultado_consulta_{timestamp}.xlsx"
        
        print_with_timestamp(f"Salvando resultados em {output_file}...")
        
        # Salva os resultados em uma planilha Excel com múltiplas abas
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # Aba 1: Comparação Detalhada
            #df1.to_excel(writer, sheet_name='Comparação Detalhada', index=False)
            
            # Aba 2: Agrupado por Código Esocial
            df2.to_excel(writer, sheet_name='Agrupado por Código Esocial', index=False)
            
            # Aba 3: Eventos Não Mapeados
            df3.to_excel(writer, sheet_name='Agrupado por Descr. Esocial', index=False)
            
            # Aba 4: Resultado da tabela esocial
            df4.to_excel(writer, sheet_name='Resultado da tabela esocial', index=False)
            
            # Ajusta a largura das colunas em todas as abas
            for sheet_name in writer.sheets:
                worksheet = writer.sheets[sheet_name]
                df_to_use = None
                
                #if sheet_name == 'Comparação Detalhada':
                #    df_to_use = df1
                if sheet_name == 'Agrupado por Código Esocial':
                    df_to_use = df2
                elif sheet_name == 'Agrupado por Descr. Esocial':
                    df_to_use = df3
                elif sheet_name == 'Resultado da tabela esocial':
                    df_to_use = df4
                
                for idx, col in enumerate(df_to_use.columns):
                    max_length = max(
                        df_to_use[col].astype(str).apply(len).max(),
                        len(str(col))
                    )
                    worksheet.column_dimensions[get_column_letter(idx + 1)].width = max_length + 2
                    
                    # Formata as colunas de valores como número com duas casas decimais
                    if 'valor' in col.lower() or 'total' in col.lower():
                        for cell in worksheet[get_column_letter(idx + 1)][1:]:
                            cell.number_format = '#,##0.00'
                
                # Adiciona formatação condicional para a aba 'Agrupado por Código Esocial'
                if sheet_name == 'Agrupado por Código Esocial':
                    # Encontra o índice da coluna Resultado_Final
                    resultado_final_idx = None
                    for idx, col in enumerate(df2.columns):
                        if col == 'Resultado_Final':
                            resultado_final_idx = idx
                            break
                    
                    if resultado_final_idx is not None:
                        # Define a cor verde para linhas com Resultado_Final = 0
                        verde = PatternFill(start_color='90EE90', end_color='90EE90', fill_type='solid')
                        
                        # Aplica a formatação condicional
                        for row_idx, row in enumerate(worksheet.iter_rows(min_row=2), start=2):
                            valor_resultado = row[resultado_final_idx].value
                            if valor_resultado is not None:
                                try:
                                    # Verifica se o valor é próximo de zero (para lidar com erros de arredondamento)
                                    if abs(float(valor_resultado)) < 0.01:
                                        # Aplica a cor verde a toda a linha
                                        for cell in row:
                                            cell.fill = verde
                                except (ValueError, TypeError):
                                    # Ignora valores não numéricos
                                    pass
                
                # Adiciona formatação condicional para a aba 'Agrupado por Descr. Esocial'
                if sheet_name == 'Agrupado por Descr. Esocial':
                    # Encontra o índice da coluna Resultado_Final
                    resultado_final_idx = None
                    for idx, col in enumerate(df3.columns):
                        if col == 'Resultado_Final':
                            resultado_final_idx = idx
                            break
                    
                    if resultado_final_idx is not None:
                        # Define a cor verde para linhas com Resultado_Final = 0
                        verde = PatternFill(start_color='90EE90', end_color='90EE90', fill_type='solid')
                        
                        # Aplica a formatação condicional
                        for row_idx, row in enumerate(worksheet.iter_rows(min_row=2), start=2):
                            valor_resultado = row[resultado_final_idx].value
                            if valor_resultado is not None:
                                try:
                                    # Verifica se o valor é próximo de zero (para lidar com erros de arredondamento)
                                    if abs(float(valor_resultado)) < 0.01:
                                        # Aplica a cor verde a toda a linha
                                        for cell in row:
                                            cell.fill = verde
                                except (ValueError, TypeError):
                                    # Ignora valores não numéricos
                                    pass
        
        print_with_timestamp(f"\nResultados salvos em: {output_file}")
        
        conn.close()
        print_with_timestamp("\nConsultas concluídas com sucesso!")
        
    except Exception as e:
        print_with_timestamp(f"Erro ao executar consultas: {str(e)}")

if __name__ == "__main__":
    verificar_dados()
    consultar_dados_join()