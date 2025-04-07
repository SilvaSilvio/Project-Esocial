import pandas as pd
from read_excel import transform_data

def group_and_sum(df):
    """
    Agrupa os registros iguais e soma os valores
    """
    # Define as colunas para agrupamento (todas exceto Valor)
    group_columns = [col for col in df.columns if col not in ['Valor']]
    
    # Agrupa e soma os valores
    df_grouped = df.groupby(group_columns)['Valor'].sum().reset_index()
    
    # Ordena o resultado
    df_grouped = df_grouped.sort_values(by=['Código', 'Valor'], na_position='last')
    
    return df_grouped

# Teste com o arquivo original
file_path = "ANDREANI_dirf x esocial 1.xlsx"
df = pd.read_excel(file_path)
df_transformed = transform_data(df)
df_grouped = group_and_sum(df_transformed)

# Mostra algumas linhas do resultado agrupado:
print("\nPrimeiras 5 linhas do resultado agrupado:")
print(df_grouped.head())

# Mostra algumas estatísticas
print("\nEstatísticas:")
print(f"Total de linhas originais: {len(df_transformed)}")
print(f"Total de linhas após agrupamento: {len(df_grouped)}")
print(f"Redução de registros: {((len(df_transformed) - len(df_grouped)) / len(df_transformed) * 100):.2f}%")

# Mostra alguns exemplos de agrupamento
print("\nExemplos de agrupamento:")
for _, row in df_grouped.head().iterrows():
    print(f"\nCódigo: {row['Código']}")
    print(f"Descrição: {row['Descrição']}")
    print(f"Valor Total: {row['Valor']:.2f}") 