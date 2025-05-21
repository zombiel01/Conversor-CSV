# Conversor de XLS para CSV

Este script Python converte arquivos Excel (XLS/XLSX) para o formato CSV.

## Requisitos

- Python 3.x
- Bibliotecas: pandas, openpyxl

Para instalar as bibliotecas necessárias:
```
pip install pandas openpyxl
```

## Como usar

### Uso básico
```
python xls_to_csv.py arquivo_entrada.xls
```

Isso irá converter o arquivo Excel para CSV usando o mesmo nome de arquivo, mas com extensão .csv.

### Especificando o arquivo de saída
```
python xls_to_csv.py arquivo_entrada.xls arquivo_saida.csv
```

### Especificando uma planilha específica
```
python xls_to_csv.py arquivo_entrada.xls arquivo_saida.csv 2
```
Onde o número 2 representa a terceira planilha (a contagem começa em 0).

## Funcionalidades

- Converte arquivos XLS/XLSX para CSV
- Permite especificar o arquivo de saída
- Permite selecionar uma planilha específica
- Codificação UTF-8 para suporte a caracteres especiais
- Tratamento de erros para arquivos inexistentes ou formatos inválidos

## Exemplo

Arquivo de exemplo incluído:
- `teste.xlsx`: Arquivo Excel de exemplo
- `teste.csv`: Resultado da conversão

## Observações

- O script usa a codificação UTF-8 para o arquivo CSV de saída
- Os índices das linhas não são incluídos no arquivo CSV
