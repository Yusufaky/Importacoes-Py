import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
import openpyxl
import pyodbc
import os
import numpy as np


def processar_arquivo_GALP():

    # Abrir a caixa de diálogo de seleção de arquivo
    filename = filedialog.askopenfilename(
        initialdir='/', title='Selecione o arquivo', filetypes=[('Arquivos do Excel', '*.xlsx')])
    if (filename == ''):
        messagebox.showinfo('Erro Sem Ficheiro',
                            'Nenhum arquivo foi selecionado.')
    else:
        # Remove o caminho e a extensao do nome do ficheiro
        nome_arquivo, extensao = os.path.splitext(os.path.basename(filename))

        # Carregando o arquivo XLSX
        workbook = openpyxl.load_workbook(filename)
        # Salvando o arquivo como XLS
        workbook.save('C:\\python\\' + nome_arquivo + '.xls')

        # Exibir uma mensagem de conclusão
        messagebox.showinfo(
            'Concluído', 'O arquivo foi processado com sucesso.')


def processar_arquivo_REPSOL():
    # Abrir a caixa de diálogo de seleção de arquivo
    filename = filedialog.askopenfilename(
        initialdir='/', title='Selecione o arquivo', filetypes=[('Arquivos do Excel', '*.xlsx')])
    if (filename == ''):
        messagebox.showinfo('Erro Sem Ficheiro',
                            'Nenhum arquivo foi selecionado.')
    else:

        # Carregar o arquivo Excel em um DataFrame
        df = pd.read_excel(filename)
        # Remove o caminho e a extensao do nome do ficheiro
        nome_arquivo, extensao = os.path.splitext(os.path.basename(filename))

        cabecalho = list(df)
        # Calcular os Valores das colunas
        df['Sem_IVA'] = df['IMP_TOTAL'] / (1+(df['IVA']/100))

        # Obtemos a posição do PAIS_DE_SERVICO
        pos = cabecalho.index('IMP_TOTAL')

        # Adiciona a coluna Sem_IVA ao cabeçalho
        cabecalho.insert(pos + 1, 'Sem_IVA')

        # Selecionar as colunas a serem exportadas
        colunas = cabecalho
        # Exportar o DataFrame para um arquivo XLSX com as colunas selecionadas
        df[colunas].to_excel(
            'C:\\python\\' + nome_arquivo + '.xlsx', index=False)

        # Carregando o arquivo XLSX
        workbook = openpyxl.load_workbook(
            'C:\\python\\' + nome_arquivo + '.xlsx')

        # Salvando o arquivo como XLS
        workbook.save('C:\\python\\' + nome_arquivo + '.xls')

        # Remove o arquivo em XLSX
        os.remove('C:\\python\\' + nome_arquivo + '.xlsx')

        # Exibir uma mensagem de conclusão
        messagebox.showinfo(
            'Concluído', 'O arquivo foi processado com sucesso.')


def processar_arquivo_VALCARCE():

    # Abrir a caixa de diálogo de seleção de arquivo
    filename = filedialog.askopenfilename(
        initialdir='/', title='Selecione o arquivo', filetypes=[('Arquivos do Excel', '*.txt')])
    if (filename == ''):
        messagebox.showinfo('Erro Sem Ficheiro',
                            'Nenhum arquivo foi selecionado.')
    else:
        # Carregar o arquivo Excel em um DataFrame
        df = pd.read_csv(filename, delimiter='\t')
        # Remove o caminho e a extensao do nome do ficheiro
        nome_arquivo, extensao = os.path.splitext(os.path.basename(filename))

        # Exportar o DataFrame para um arquivo XLSX com as colunas selecionadas
        df.to_excel('C:\\python\\' + nome_arquivo + '.xlsx', index=False)

        ficheiro = 'C:\\python\\' + nome_arquivo + '.xlsx'
        # Carregando o arquivo XLSX
        df = pd.read_excel(ficheiro)
        cabecalho = list(df)
        #  Altera as "," para "." para converter em float
        Cantidad = [s.replace(",", ".") for s in df['Cantidad']]
        Precio = [s.replace(",", ".") for s in df['Precio']]
        Dto = [s.replace(",", ".") for s in df['Dto.Clien']]

        # Converte as colunas em float
        valores_Cantidad = []
        for valor in Cantidad:
            valor_float = float(valor)
            valores_Cantidad.append(valor_float)

        # Converte as colunas em float
        valores_Precio = []
        for valor in Precio:
            valor_float = float(valor)
            valores_Precio.append(valor_float)

        # Calcula os totais
        totais = []
        for i in range(len(valores_Precio)):
            total = valores_Precio[i] * valores_Cantidad[i]
            totais.append(total)

        df['TOTA_S/IVA'] = totais

        # Converte as colunas em float
        valores_Dto = []
        for valor in Dto:
            valor_float = float(valor)
            valores_Dto.append(valor_float)

        # Calcula os totais
        VALOR_UNITARIO = []
        for i in range(len(valores_Precio)):
            total = ((valores_Precio[i] - valores_Dto[i])/1.21)
            VALOR_UNITARIO.append(total)

        df['VALOR_UNITARIO'] = VALOR_UNITARIO

        # Seleciona a os elementos do Cabeçalho

        # Adiciona a coluna VALOR_UNITARIO e TOTA_S/IVA ao cabeçalho
        cabecalho.append('TOTA_S/IVA')
        cabecalho.append('VALOR_UNITARIO')

        # Selecionar as colunas a serem exportadas
        colunas = cabecalho

        # Exportar o DataFrame para um arquivo XLSX com as colunas selecionadas
        df[colunas].to_excel(
            'C:\\python\\' + nome_arquivo + '.xlsx', index=False)

        # Exibir uma mensagem de conclusão
        messagebox.showinfo(
            'Concluído', 'O arquivo foi processado com sucesso.')


def processar_arquivo_VIAVERDE():

    # Abrir a caixa de diálogo de seleção de arquivo
    filename = filedialog.askopenfilename(
        initialdir='/', title='Selecione o arquivo', filetypes=[('Arquivos do Excel', '*.xlsx')])
    if (filename == ''):
        messagebox.showinfo('Erro Sem Ficheiro',
                            'Nenhum arquivo foi selecionado.')
    else:
        # Carregar o arquivo Excel em um DataFrame
        df = pd.read_excel(filename, header=None)
        nome_arquivo, extensao = os.path.splitext(os.path.basename(filename))

        cabecalho = df.iloc[0]   # obtém a primeira linha do dataframe

        if (cabecalho[0] == 'License Plate'):
            df = df.drop([0])   # elimina a primeira linha do dataframe
        # escreve o dataframe sem cabeçalho em um novo arquivo Excel
        df.to_excel('C:\\python\\' + nome_arquivo +
                    '.xlsx', index=False, header=None)

        # Exibir uma mensagem de conclusão
        messagebox.showinfo(
            'Concluído', 'O arquivo foi processado com sucesso.')


def processar_arquivo_VIALTIS():

    # Abrir a caixa de diálogo de seleção de arquivo
    filename = filedialog.askopenfilename(
        initialdir='/', title='Selecione o arquivo', filetypes=[('Arquivos do Excel', '*.xlsx')])
    if (filename == ''):
        messagebox.showinfo('Erro Sem Ficheiro',
                            'Nenhum arquivo foi selecionado.')
    else:
        # Carregar o arquivo Excel em um DataFrame
        df = pd.read_excel(filename)
        # Remove o caminho e a extensao do nome do ficheiro
        nome_arquivo, extensao = os.path.splitext(os.path.basename(filename))

        cabecalho = list(df)
        # Calcular os Valores das colunas
        df['PAIS_DE_SERVICO2'] = df['PAIS_DE_SERVICO']

        # Obtemos a posição do PAIS_DE_SERVICO
        pos = cabecalho.index('PAIS_DE_SERVICO')

        # Adiciona a coluna PAIS_DE_SERVICO2 ao cabeçalho
        cabecalho.insert(pos + 1, 'PAIS_DE_SERVICO2')

        # Selecionar as colunas a serem exportadas
        colunas = cabecalho
        # Colunas a serem exportadas AINDA ESTATICO

        # Exportar o DataFrame para um arquivo XLSX com as colunas selecionadas
        df[colunas].to_excel(
            'C:\\python\\' + nome_arquivo + '.xlsx', index=False)

        # Exibir uma mensagem de conclusão
        messagebox.showinfo(
            'Concluído', 'O arquivo foi processado com sucesso.')


# Criar a janela principal
root = tk.Tk()

# Definir o tamanho da janela
root.geometry('800x600')

# Definir o título da janela
root.title('Carregamento de Arquivo Excel')

# Definir a cor de fundo da janela
root.configure(bg='white')

# Criar um botão "Selecionar arquivo"
botao_selecionarGALP = tk.Button(root, text='Selecionar arquivo GALP',
                                 command=processar_arquivo_GALP, bg='blue', fg='white', font=('Arial', 12, 'bold'))
botao_selecionarGALP.pack(pady=35)

botao_selecionarREPSOL = tk.Button(root, text='Selecionar arquivo REPSOL',
                                   command=processar_arquivo_REPSOL, bg='blue', fg='white', font=('Arial', 12, 'bold'))
botao_selecionarREPSOL.pack(pady=35)

botao_selecionarVALCARCER = tk.Button(root, text='Selecionar arquivo VALCARCER',
                                      command=processar_arquivo_VALCARCE, bg='blue', fg='white', font=('Arial', 12, 'bold'))
botao_selecionarVALCARCER.pack(pady=35)

botao_selecionarVIAVERDE = tk.Button(root, text='Selecionar arquivo VIA VERDE',
                                     command=processar_arquivo_VIAVERDE, bg='blue', fg='white', font=('Arial', 12, 'bold'))
botao_selecionarVIAVERDE.pack(pady=35)

botao_selecionarVIALTIS = tk.Button(root, text='Selecionar arquivo VIALTIS',
                                    command=processar_arquivo_VIALTIS, bg='blue', fg='white', font=('Arial', 12, 'bold'))
botao_selecionarVIALTIS.pack(pady=35)

# Iniciar a janela
root.mainloop()
