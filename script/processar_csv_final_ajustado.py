#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import csv
import re
import sys
import glob
from datetime import datetime

def extrair_valores_totais(linhas):
    """Extrai e soma todos os valores monetários de linhas que contenham 'Total'"""
    soma_valores = 0.0
    
    # Verificar as últimas 10 linhas do arquivo (ou todas, se tiver menos de 10)
    linhas_verificar = min(10, len(linhas))
    for i in range(len(linhas) - linhas_verificar, len(linhas)):
        if i >= 0 and "Total" in linhas[i]:
            partes = linhas[i].split(';')
            
            # Procura por valores após qualquer padrão "Total X:"
            for j, parte in enumerate(partes):
                if "Total" in parte and ":" in parte and j + 1 < len(partes):
                    # Ignora se for "Total Procedimentos:"
                    if "Total Procedimentos:" == parte.strip():
                        continue
                    
                    # Pega o valor e converte para formato numérico
                    valor_str = partes[j + 1].strip().replace('.', '').replace(',', '.')
                    # Remove aspas ou outros caracteres não numéricos
                    valor_str = re.sub(r'[^\d.]', '', valor_str)
                    if valor_str:
                        try:
                            valor = float(valor_str)
                            print(f"  - Valor encontrado: {valor} (linha: {linhas[i].strip()})")
                            soma_valores += valor
                        except ValueError:
                            print(f"  - Erro ao converter valor: {partes[j + 1]}")
    
    return soma_valores

def processar_arquivos_csv(diretorio, padrao_arquivos, arquivo_saida):
    """Processa múltiplos arquivos CSV e gera um arquivo consolidado"""
    arquivos = glob.glob(os.path.join(diretorio, padrao_arquivos))
    
    if not arquivos:
        print(f"Nenhum arquivo encontrado com o padrão: {padrao_arquivos}")
        return
    
    # Ordenar arquivos pelo nome para processamento consistente
    arquivos.sort()
    
    # Inicializar variáveis
    linhas_consolidadas = []
    soma_total = 0.0
    cabecalhos = []
    
    # Processar cada arquivo
    for arquivo in arquivos:
        print(f"Processando arquivo: {arquivo}")
        
        with open(arquivo, 'r', encoding='utf-8-sig') as f:
            linhas = f.readlines()
            
            # Guardar os cabeçalhos do primeiro arquivo
            if not cabecalhos and len(linhas) >= 3:
                cabecalhos = linhas[:3]
            
            # Extrair linhas a partir da quarta linha
            if len(linhas) >= 4:
                dados = linhas[3:]
                
                # Extrair e somar todos os valores de total
                valor = extrair_valores_totais(dados)
                soma_total += valor
                print(f"Valor total extraído do arquivo {os.path.basename(arquivo)}: {valor}, Soma acumulada: {soma_total}")
                
                # Adicionar linhas de dados ao consolidado
                linhas_consolidadas.extend(dados)
    
    # Criar arquivo de saída
    with open(arquivo_saida, 'w', encoding='utf-8') as f_saida:
        # Escrever cabeçalhos
        f_saida.writelines(cabecalhos)
        
        # Escrever linhas consolidadas
        f_saida.writelines(linhas_consolidadas)
        
        # Adicionar linha com a soma total
        valor_formatado = f"{soma_total:.2f}".replace('.', ',')
        f_saida.write(f"\nTotal Geral de Todos os Arquivos:;;Total:;{valor_formatado};\n")
    
    print(f"\nProcessamento concluído!")
    print(f"Arquivo consolidado gerado: {arquivo_saida}")
    print(f"Soma total de todos os arquivos: {valor_formatado}")

if __name__ == "__main__":
    # Verificar argumentos da linha de comando
    if len(sys.argv) < 2:
        diretorio = input("Digite o diretório dos arquivos CSV (ou pressione Enter para usar o diretório atual): ").strip()
        if not diretorio:
            diretorio = os.getcwd()
        
        padrao_arquivos = input("Digite o padrão dos arquivos a serem processados (ex: 'anne colono *.csv'): ").strip()
        if not padrao_arquivos:
            padrao_arquivos = "*.csv"
        
        arquivo_saida = input("Digite o nome do arquivo de saída (ou pressione Enter para usar 'consolidado.csv'): ").strip()
        if not arquivo_saida:
            arquivo_saida = "consolidado.csv"
    else:
        diretorio = sys.argv[1] if len(sys.argv) > 1 else os.getcwd()
        padrao_arquivos = sys.argv[2] if len(sys.argv) > 2 else "*.csv"
        arquivo_saida = sys.argv[3] if len(sys.argv) > 3 else "consolidado.csv"
    
    # Garantir que o arquivo de saída tenha caminho absoluto
    if not os.path.isabs(arquivo_saida):
        arquivo_saida = os.path.join(diretorio, arquivo_saida)
    
    # Processar os arquivos
    processar_arquivos_csv(diretorio, padrao_arquivos, arquivo_saida)
