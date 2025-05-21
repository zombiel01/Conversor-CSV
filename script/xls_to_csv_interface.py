#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Script para converter arquivos XLS/XLSX para CSV.
Uso: python3 xls_to_csv.py [arquivo_entrada.xls] [arquivo_saida.csv] [numero_planilha]

Se executado sem argumentos, o script entrará no modo interativo.
Se o arquivo de saída não for especificado, será usado o mesmo nome do arquivo de entrada com extensão .csv
Se o número da planilha não for especificado, será usada a primeira planilha (índice 0)
"""

import sys
import os
import pandas as pd
import glob

def listar_arquivos_excel(diretorio='.'):
    """
    Lista todos os arquivos Excel no diretório especificado.
    
    Args:
        diretorio (str): Diretório para buscar arquivos Excel
    
    Returns:
        list: Lista de arquivos Excel encontrados
    """
    arquivos_xls = glob.glob(os.path.join(diretorio, '*.xls'))
    arquivos_xlsx = glob.glob(os.path.join(diretorio, '*.xlsx'))
    return sorted(arquivos_xls + arquivos_xlsx)

def listar_planilhas(arquivo_excel):
    """
    Lista todas as planilhas disponíveis em um arquivo Excel.
    
    Args:
        arquivo_excel (str): Caminho para o arquivo Excel
    
    Returns:
        list: Lista de nomes das planilhas
    """
    try:
        xl = pd.ExcelFile(arquivo_excel)
        return xl.sheet_names
    except Exception as e:
        print(f"Erro ao listar planilhas: {str(e)}")
        return []

def converter_xls_para_csv(arquivo_entrada, arquivo_saida=None, numero_planilha=0):
    """
    Converte um arquivo XLS/XLSX para CSV.
    
    Args:
        arquivo_entrada (str): Caminho para o arquivo XLS/XLSX de entrada
        arquivo_saida (str, opcional): Caminho para o arquivo CSV de saída
        numero_planilha (int, opcional): Número da planilha a ser convertida (padrão: 0)
    
    Returns:
        bool: True se a conversão foi bem-sucedida, False caso contrário
    """
    try:
        # Verifica se o arquivo de entrada existe
        if not os.path.isfile(arquivo_entrada):
            print(f"Erro: O arquivo '{arquivo_entrada}' não existe.")
            return False
        
        # Se o arquivo de saída não for especificado, usa o mesmo nome do arquivo de entrada com extensão .csv
        if arquivo_saida is None:
            nome_base = os.path.splitext(arquivo_entrada)[0]
            arquivo_saida = f"{nome_base}.csv"
        
        # Lê o arquivo XLS/XLSX
        print(f"Lendo o arquivo '{arquivo_entrada}'...")
        
        # Obtém os nomes das planilhas
        planilhas = listar_planilhas(arquivo_entrada)
        if not planilhas:
            print("Erro: Não foi possível ler as planilhas do arquivo.")
            return False
        
        # Verifica se o número da planilha é válido
        if isinstance(numero_planilha, int):
            if numero_planilha < 0 or numero_planilha >= len(planilhas):
                print(f"Erro: Número de planilha inválido. O arquivo tem {len(planilhas)} planilha(s).")
                return False
            sheet_name = numero_planilha
        else:
            # Se for um nome de planilha
            if numero_planilha not in planilhas:
                print(f"Erro: Planilha '{numero_planilha}' não encontrada no arquivo.")
                return False
            sheet_name = numero_planilha
        
        # Lê a planilha especificada
        df = pd.read_excel(arquivo_entrada, sheet_name=sheet_name)
        
        # Salva como CSV
        print(f"Convertendo para CSV e salvando como '{arquivo_saida}'...")
        df.to_csv(arquivo_saida, index=False, encoding='utf-8')
        
        print(f"Conversão concluída com sucesso! O arquivo foi salvo como '{arquivo_saida}'.")
        return True
    
    except Exception as e:
        print(f"Erro durante a conversão: {str(e)}")
        return False

def modo_interativo():
    """
    Executa o script no modo interativo, solicitando informações ao usuário.
    """
    print("\n===== CONVERSOR DE XLS/XLSX PARA CSV =====\n")
    
    # Lista arquivos Excel no diretório atual
    arquivos_excel = listar_arquivos_excel()
    
    if not arquivos_excel:
        diretorio = input("Nenhum arquivo Excel encontrado no diretório atual.\nInforme o caminho completo do arquivo Excel: ")
        if os.path.isdir(diretorio):
            arquivos_excel = listar_arquivos_excel(diretorio)
            if not arquivos_excel:
                print(f"Nenhum arquivo Excel encontrado em '{diretorio}'.")
                arquivo_entrada = input("Informe o caminho completo do arquivo Excel: ")
            else:
                print("\nArquivos Excel disponíveis:")
                for i, arquivo in enumerate(arquivos_excel):
                    print(f"{i+1}. {arquivo}")
                
                while True:
                    try:
                        escolha = int(input("\nEscolha o número do arquivo (ou 0 para informar outro caminho): "))
                        if escolha == 0:
                            arquivo_entrada = input("Informe o caminho completo do arquivo Excel: ")
                            break
                        elif 1 <= escolha <= len(arquivos_excel):
                            arquivo_entrada = arquivos_excel[escolha-1]
                            break
                        else:
                            print("Opção inválida. Tente novamente.")
                    except ValueError:
                        print("Por favor, digite um número válido.")
        else:
            arquivo_entrada = diretorio
    else:
        print("Arquivos Excel disponíveis no diretório atual:")
        for i, arquivo in enumerate(arquivos_excel):
            print(f"{i+1}. {arquivo}")
        
        while True:
            try:
                escolha = int(input("\nEscolha o número do arquivo (ou 0 para informar outro caminho): "))
                if escolha == 0:
                    arquivo_entrada = input("Informe o caminho completo do arquivo Excel: ")
                    break
                elif 1 <= escolha <= len(arquivos_excel):
                    arquivo_entrada = arquivos_excel[escolha-1]
                    break
                else:
                    print("Opção inválida. Tente novamente.")
            except ValueError:
                print("Por favor, digite um número válido.")
    
    # Verifica se o arquivo existe
    if not os.path.isfile(arquivo_entrada):
        print(f"Erro: O arquivo '{arquivo_entrada}' não existe.")
        return False
    
    # Lista as planilhas disponíveis
    planilhas = listar_planilhas(arquivo_entrada)
    if not planilhas:
        print("Erro: Não foi possível ler as planilhas do arquivo.")
        return False
    
    print("\nPlanilhas disponíveis:")
    for i, planilha in enumerate(planilhas):
        print(f"{i+1}. {planilha}")
    
    # Solicita a escolha da planilha
    while True:
        try:
            escolha_planilha = input("\nEscolha o número da planilha (ou pressione Enter para usar a primeira): ")
            if escolha_planilha == "":
                numero_planilha = 0
                break
            else:
                escolha_planilha = int(escolha_planilha)
                if 1 <= escolha_planilha <= len(planilhas):
                    numero_planilha = escolha_planilha - 1
                    break
                else:
                    print("Opção inválida. Tente novamente.")
        except ValueError:
            print("Por favor, digite um número válido.")
    
    # Solicita o nome do arquivo de saída
    nome_base = os.path.splitext(os.path.basename(arquivo_entrada))[0]
    sugestao_saida = f"{nome_base}.csv"
    
    arquivo_saida = input(f"\nNome do arquivo CSV de saída (ou pressione Enter para usar '{sugestao_saida}'): ")
    if arquivo_saida == "":
        arquivo_saida = sugestao_saida
    
    # Confirma as escolhas
    print("\nResumo da conversão:")
    print(f"Arquivo de entrada: {arquivo_entrada}")
    print(f"Planilha: {planilhas[numero_planilha]} (índice {numero_planilha})")
    print(f"Arquivo de saída: {arquivo_saida}")
    
    confirmacao = input("\nConfirmar conversão? (s/n): ")
    if confirmacao.lower() not in ['s', 'sim', 'y', 'yes']:
        print("Conversão cancelada pelo usuário.")
        return False
    
    # Realiza a conversão
    return converter_xls_para_csv(arquivo_entrada, arquivo_saida, numero_planilha)

def main():
    # Verifica se há argumentos da linha de comando
    if len(sys.argv) == 1:
        # Sem argumentos, entra no modo interativo
        sucesso = modo_interativo()
    else:
        # Com argumentos, usa o modo tradicional
        # Obtém o arquivo de entrada
        arquivo_entrada = sys.argv[1]
        
        # Obtém o arquivo de saída, se especificado
        arquivo_saida = None
        if len(sys.argv) >= 3:
            arquivo_saida = sys.argv[2]
        
        # Obtém o número da planilha, se especificado
        numero_planilha = 0
        if len(sys.argv) >= 4:
            try:
                numero_planilha = int(sys.argv[3])
            except ValueError:
                print("Erro: O número da planilha deve ser um número inteiro.")
                sys.exit(1)
        
        # Converte o arquivo
        sucesso = converter_xls_para_csv(arquivo_entrada, arquivo_saida, numero_planilha)
    
    # Sai com código de erro apropriado
    sys.exit(0 if sucesso else 1)

if __name__ == "__main__":
    main()
