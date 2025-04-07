# Processamento de Dados eSocial e LG

Este projeto contém scripts Python para processar e comparar dados do eSocial com dados da Ficha Financeira (LG).

## Requisitos

- Python 3.8 ou superior
- Dependências listadas em `requirements.txt`

## Instalação

1. Clone este repositório
2. Instale as dependências:
```bash
pip install -r requirements.txt
```

## Estrutura de Arquivos

- `main.py`: Script principal que integra o processamento do eSocial e LG
- `process_esocial.py`: Módulo para processamento dos dados do eSocial
- `process_lg_depara.py`: Módulo para processamento dos dados da Ficha Financeira e De/Para
- `requirements.txt`: Lista de dependências Python

## Arquivos de Entrada Necessários

1. Arquivo do eSocial (formato Excel)
2. Arquivo da Ficha Financeira (formato CSV ou Excel)
3. Arquivo De/Para (formato CSV ou Excel)

## Como Usar

1. Coloque os arquivos de entrada na mesma pasta dos scripts
2. Execute o script principal passando os nomes dos arquivos como argumentos:
```bash
python main.py <arquivo_esocial> <arquivo_ficha_financeira> <arquivo_depara>
```

Exemplo:
```bash
python main.py ANDREANI_dirf_x_esocial_1.xlsx ficha_financeira.csv depara.csv
```

O script irá:
1. Verificar se os arquivos existem
2. Processar o arquivo do eSocial
3. Processar os arquivos da Ficha Financeira e De/Para
4. Gerar um arquivo Excel com os resultados em diferentes abas

## Solução de Problemas

### Erro de tipos de dados incompatíveis

Se você encontrar um erro como:
```
You are trying to merge on object and int64 columns for key 'CÓDIGO EVENTO'
```

Isso significa que as colunas de junção nos arquivos LG e Depara têm tipos de dados diferentes. O script foi atualizado para converter automaticamente essas colunas para o mesmo tipo (string), mas se o problema persistir, verifique se os valores nas colunas de junção são consistentes entre os arquivos.

## Saída

O script gera um arquivo Excel com o nome `resultado_comparacao_YYYYMMDD_HHMMSS.xlsx` contendo:
- Aba "eSocial": Dados processados do eSocial
- Aba "LG_Depara": Dados processados da Ficha Financeira com mapeamento do De/Para

## Suporte

Em caso de dúvidas ou problemas, por favor abra uma issue no repositório. 