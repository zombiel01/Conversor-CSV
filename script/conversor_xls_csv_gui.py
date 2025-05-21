#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Conversor de XLS/XLSX para CSV com Interface Gráfica
Este script fornece uma interface gráfica amigável para converter arquivos Excel para CSV.
"""

import os
import sys
import pandas as pd
import PySimpleGUI as sg
import threading
import traceback

# Configuração de tema e aparência
sg.theme('LightBlue2')
FONT = ('Arial', 11)
FONT_BOLD = ('Arial', 11, 'bold')

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
        sg.popup_error(f"Erro ao listar planilhas: {str(e)}")
        return []

def converter_xls_para_csv(arquivo_entrada, arquivo_saida, nome_planilha, janela):
    """
    Converte um arquivo XLS/XLSX para CSV.
    
    Args:
        arquivo_entrada (str): Caminho para o arquivo XLS/XLSX de entrada
        arquivo_saida (str): Caminho para o arquivo CSV de saída
        nome_planilha (str): Nome da planilha a ser convertida
        janela (sg.Window): Janela para atualização de progresso
    
    Returns:
        bool: True se a conversão foi bem-sucedida, False caso contrário
    """
    try:
        # Atualiza a interface
        janela.write_event_value('-PROGRESSO-', 'Lendo o arquivo Excel...')
        
        # Lê o arquivo XLS/XLSX
        df = pd.read_excel(arquivo_entrada, sheet_name=nome_planilha)
        
        # Atualiza a interface
        janela.write_event_value('-PROGRESSO-', 'Convertendo para CSV...')
        
        # Salva como CSV
        df.to_csv(arquivo_saida, index=False, encoding='utf-8')
        
        # Atualiza a interface
        janela.write_event_value('-PROGRESSO-', f'Conversão concluída com sucesso!')
        janela.write_event_value('-CONCLUIDO-', True)
        return True
    
    except Exception as e:
        # Atualiza a interface com o erro
        erro = f"Erro durante a conversão: {str(e)}"
        janela.write_event_value('-ERRO-', erro)
        traceback.print_exc()
        return False

def criar_layout_principal():
    """
    Cria o layout principal da interface gráfica.
    
    Returns:
        list: Layout para PySimpleGUI
    """
    layout = [
        [sg.Text('Conversor de Excel para CSV', font=('Arial', 16, 'bold'), justification='center', expand_x=True)],
        [sg.HorizontalSeparator()],
        
        # Seção de arquivo de entrada
        [sg.Frame('Arquivo Excel de Entrada', [
            [sg.Input(key='-ARQUIVO_ENTRADA-', size=(50, 1), readonly=True),
             sg.FileBrowse('Selecionar Arquivo', file_types=(("Arquivos Excel", "*.xls;*.xlsx"),), key='-BROWSE-')]
        ])],
        
        # Seção de seleção de planilha
        [sg.Frame('Planilha', [
            [sg.Text('Selecione a planilha:')],
            [sg.Combo(values=[], key='-PLANILHA-', size=(48, 1), readonly=True, disabled=True),
             sg.Button('Carregar Planilhas', key='-CARREGAR_PLANILHAS-', disabled=True)]
        ])],
        
        # Seção de arquivo de saída
        [sg.Frame('Arquivo CSV de Saída', [
            [sg.Input(key='-ARQUIVO_SAIDA-', size=(50, 1), readonly=True),
             sg.SaveAs('Selecionar Destino', file_types=(("Arquivo CSV", "*.csv"),), key='-SAVE_AS-', disabled=True)]
        ])],
        
        # Barra de progresso e status
        [sg.Text('Status:', font=FONT_BOLD), sg.Text('Aguardando seleção de arquivo...', key='-STATUS-', size=(50, 1))],
        [sg.ProgressBar(100, orientation='h', size=(44, 20), key='-PROGRESS_BAR-', visible=False)],
        
        # Botões de ação
        [sg.Button('Converter', key='-CONVERTER-', disabled=True), 
         sg.Button('Cancelar', key='-CANCELAR-'),
         sg.Button('Sobre', key='-SOBRE-')]
    ]
    
    return layout

def main():
    """
    Função principal que cria e gerencia a interface gráfica.
    """
    # Cria a janela principal
    janela = sg.Window('Conversor XLS para CSV', criar_layout_principal(), finalize=True, icon=sg.EMOJI_BASE64_HAPPY_JOY)
    
    # Variáveis de controle
    thread_conversao = None
    conversao_em_andamento = False
    
    # Loop de eventos
    while True:
        evento, valores = janela.read(timeout=100)
        
        # Verifica se a janela foi fechada
        if evento == sg.WIN_CLOSED or evento == '-CANCELAR-':
            if conversao_em_andamento:
                sg.popup_yes_no('Uma conversão está em andamento. Deseja realmente sair?', title='Confirmar Saída')
            break
        
        # Quando um arquivo de entrada é selecionado
        if evento == '-BROWSE-' and valores['-ARQUIVO_ENTRADA-']:
            # Habilita o botão de carregar planilhas
            janela['-CARREGAR_PLANILHAS-'].update(disabled=False)
            
            # Sugere um nome para o arquivo de saída
            nome_base = os.path.splitext(os.path.basename(valores['-ARQUIVO_ENTRADA-']))[0]
            sugestao_saida = f"{nome_base}.csv"
            janela['-ARQUIVO_SAIDA-'].update(sugestao_saida)
            janela['-SAVE_AS-'].update(disabled=False)
            
            # Atualiza o status
            janela['-STATUS-'].update(f"Arquivo selecionado: {os.path.basename(valores['-ARQUIVO_ENTRADA-'])}")
        
        # Quando o botão de carregar planilhas é clicado
        if evento == '-CARREGAR_PLANILHAS-' and valores['-ARQUIVO_ENTRADA-']:
            try:
                # Lista as planilhas do arquivo
                planilhas = listar_planilhas(valores['-ARQUIVO_ENTRADA-'])
                
                if planilhas:
                    # Atualiza o combo de planilhas
                    janela['-PLANILHA-'].update(values=planilhas, value=planilhas[0], disabled=False)
                    
                    # Habilita o botão de converter
                    janela['-CONVERTER-'].update(disabled=False)
                    
                    # Atualiza o status
                    janela['-STATUS-'].update(f"Planilhas carregadas: {len(planilhas)} encontradas")
                else:
                    sg.popup_error("Não foi possível encontrar planilhas no arquivo selecionado.")
            except Exception as e:
                sg.popup_error(f"Erro ao carregar planilhas: {str(e)}")
        
        # Quando o botão de converter é clicado
        if evento == '-CONVERTER-' and not conversao_em_andamento:
            # Verifica se todos os campos necessários estão preenchidos
            if not valores['-ARQUIVO_ENTRADA-']:
                sg.popup_error("Selecione um arquivo Excel de entrada.")
                continue
                
            if not valores['-PLANILHA-']:
                sg.popup_error("Selecione uma planilha.")
                continue
                
            if not valores['-ARQUIVO_SAIDA-']:
                sg.popup_error("Defina o arquivo CSV de saída.")
                continue
            
            # Confirma a conversão
            if sg.popup_yes_no(
                f"Confirma a conversão?\n\n"
                f"Arquivo de entrada: {os.path.basename(valores['-ARQUIVO_ENTRADA-'])}\n"
                f"Planilha: {valores['-PLANILHA-']}\n"
                f"Arquivo de saída: {os.path.basename(valores['-ARQUIVO_SAIDA-'])}",
                title="Confirmar Conversão"
            ) == "Yes":
                # Inicia a conversão em uma thread separada
                conversao_em_andamento = True
                janela['-STATUS-'].update("Iniciando conversão...")
                janela['-PROGRESS_BAR-'].update(visible=True, current_count=10)
                janela['-CONVERTER-'].update(disabled=True)
                janela['-CARREGAR_PLANILHAS-'].update(disabled=True)
                janela['-BROWSE-'].update(disabled=True)
                janela['-SAVE_AS-'].update(disabled=True)
                
                thread_conversao = threading.Thread(
                    target=converter_xls_para_csv,
                    args=(valores['-ARQUIVO_ENTRADA-'], valores['-ARQUIVO_SAIDA-'], valores['-PLANILHA-'], janela),
                    daemon=True
                )
                thread_conversao.start()
        
        # Atualização de progresso da conversão
        if evento == '-PROGRESSO-':
            janela['-STATUS-'].update(valores[evento])
            # Atualiza a barra de progresso
            if "Lendo" in valores[evento]:
                janela['-PROGRESS_BAR-'].update(current_count=30)
            elif "Convertendo" in valores[evento]:
                janela['-PROGRESS_BAR-'].update(current_count=70)
            elif "concluída" in valores[evento]:
                janela['-PROGRESS_BAR-'].update(current_count=100)
        
        # Quando a conversão é concluída
        if evento == '-CONCLUIDO-':
            conversao_em_andamento = False
            janela['-CONVERTER-'].update(disabled=False)
            janela['-CARREGAR_PLANILHAS-'].update(disabled=False)
            janela['-BROWSE-'].update(disabled=False)
            janela['-SAVE_AS-'].update(disabled=False)
            
            # Pergunta se deseja abrir o arquivo
            if sg.popup_yes_no(
                f"Conversão concluída com sucesso!\n\n"
                f"O arquivo foi salvo como: {valores['-ARQUIVO_SAIDA-']}\n\n"
                f"Deseja converter outro arquivo?",
                title="Conversão Concluída"
            ) == "No":
                break
        
        # Quando ocorre um erro na conversão
        if evento == '-ERRO-':
            conversao_em_andamento = False
            janela['-STATUS-'].update(valores[evento])
            janela['-PROGRESS_BAR-'].update(visible=False)
            janela['-CONVERTER-'].update(disabled=False)
            janela['-CARREGAR_PLANILHAS-'].update(disabled=False)
            janela['-BROWSE-'].update(disabled=False)
            janela['-SAVE_AS-'].update(disabled=False)
            
            sg.popup_error(valores[evento])
        
        # Quando o botão Sobre é clicado
        if evento == '-SOBRE-':
            sg.popup(
                "Conversor XLS para CSV",
                "Versão 1.0",
                "Desenvolvido com PySimpleGUI e pandas",
                "Este programa converte arquivos Excel (XLS/XLSX) para o formato CSV.",
                title="Sobre o Conversor"
            )
    
    # Fecha a janela
    janela.close()

if __name__ == "__main__":
    main()
