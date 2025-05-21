# Conversor de XLS para CSV - Versão Interativa

Este script Python converte arquivos Excel (XLS/XLSX) para o formato CSV, agora com uma interface interativa de linha de comando.

## Requisitos

- Python 3.x
- Bibliotecas: pandas, openpyxl

Para instalar as bibliotecas necessárias:
```
pip install pandas openpyxl
```

## Como usar

### Modo Interativo
Simplesmente execute o script sem argumentos:
```
python xls_to_csv.py
```

O script irá:
1. Listar os arquivos Excel disponíveis no diretório atual
2. Permitir que você escolha um arquivo ou informe outro caminho
3. Mostrar as planilhas disponíveis no arquivo selecionado
4. Permitir que você escolha qual planilha converter
5. Sugerir um nome para o arquivo CSV de saída, que você pode aceitar ou alterar
6. Confirmar suas escolhas antes de realizar a conversão

### Modo com Argumentos
Você também pode usar o script com argumentos de linha de comando:
```
python xls_to_csv.py arquivo_entrada.xls [arquivo_saida.csv] [numero_planilha]
```

Onde:
- `arquivo_entrada.xls`: Caminho para o arquivo Excel a ser convertido
- `arquivo_saida.csv` (opcional): Nome do arquivo CSV de saída
- `numero_planilha` (opcional): Índice da planilha a ser convertida (começando em 0)

## Funcionalidades

- Interface interativa amigável
- Listagem automática de arquivos Excel disponíveis
- Visualização das planilhas disponíveis no arquivo
- Sugestão de nome para o arquivo de saída
- Confirmação antes da conversão
- Codificação UTF-8 para suporte a caracteres especiais
- Tratamento de erros para arquivos inexistentes ou formatos inválidos

## Exemplo

Arquivo de exemplo incluído:
- `teste.xlsx`: Arquivo Excel de exemplo
- `teste.csv`: Resultado da conversão

## Observações

- O script usa a codificação UTF-8 para o arquivo CSV de saída
- Os índices das linhas não são incluídos no arquivo CSV
