from PIL import ImageTk, Image
import datetime
import tkinter as tk
# import xlrd
# import xlwt
from tkcalendar import DateEntry
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
import pandas as pd
import openpyxl
import pyodbc
import csv
import camelot
from pdf2image import convert_from_path
import pytesseract
import os
import re
import numpy as np
import chardet
import pdfplumber
import tabula
from openpyxl import Workbook

# CRiAR O EXE
# cxfreeze seu_script.py --target-dir dist --base-name Win32GUI


def processar_arquivo_PorFazer():
    messagebox.showinfo('Falta de Configuração',
                        'A importação que escolheu ainda não foi configurada')


def processar_arquivo_Toll_Collect():
    # Abrir a caixa de diálogo de seleção de arquivo
    filename = filedialog.askopenfilename(
        initialdir='/', title='Selecione o arquivo', filetypes=[('Arquivos do Excel', '*.csv')])
    if (filename == ''):
        messagebox.showinfo('Erro Sem Ficheiro',
                            'Nenhum arquivo foi selecionado.')
    else:
        def obter_valor():
            valorFaturaEntry = entry.get()
            nome_arquivo, extensao = os.path.splitext(
                os.path.basename(filename))
            # Abrir o arquivo CSV para leitura
            with open(filename, newline='') as f_in, open('C:\\importacao\\' + nome_arquivo + '.csv', 'w', newline='') as f_out:
                # Criar um leitor CSV e um gravador CSV
                # Ler as linhas do arquivo CSV
                linhas = list(csv.reader(f_in))
                gravador = csv.writer(f_out)

                linhas = linhas[:-3]
                gravador.writerows(linhas)

            with open('C:\\importacao\\' + nome_arquivo + '.csv', 'rb') as f:
                result = chardet.detect(f.read())
                encoding = result['encoding']

            df = pd.read_csv('C:\\importacao\\' + nome_arquivo +
                             '.csv', encoding=encoding, delimiter=";")

            df.to_excel('C:\\importacao\\' + nome_arquivo +
                        '.xlsx', index=False)

            os.remove('C:\\importacao\\' + nome_arquivo + '.csv')

            df = pd.read_excel('C:\\importacao\\' + nome_arquivo +
                               '.xlsx')
            df["FATURA"] = valorFaturaEntry
            df.to_excel('C:\\importacao\\' + nome_arquivo +
                        '.xlsx', index=False)

            messagebox.showinfo(
                'Concluído', 'O arquivo foi processado com sucesso.')
            root.destroy()

        # Criar janela principal
        root = tk.Tk()
        root.resizable(width=False, height=False)
        label = tk.Label(root, text="Fatura:")
        label.pack()
        entry = tk.Entry(root)
        entry.pack()
        # Criar botão para obter o valor
        btn_obter_valor = tk.Button(
            root, text="Enviar dados", command=obter_valor)
        btn_obter_valor.pack()

        # Executar o loop principal da janela
        root.mainloop()


def processar_arquivo_NORPETROL():
    # Abrir a caixa de diálogo de seleção de arquivo
    filename = filedialog.askopenfilename(
        initialdir='/', title='Selecione o arquivo', filetypes=[('Arquivos do Excel', '*.csv')])
    if (filename == ''):
        messagebox.showinfo('Erro Sem Ficheiro',
                            'Nenhum arquivo foi selecionado.')
    else:
        def obter_valor():
            valorFaturaEntry = entry.get()
            nome_arquivo, extensao = os.path.splitext(
                os.path.basename(filename))
            with open(filename, newline='') as f_in, open('C:\\importacao\\' + nome_arquivo + '.csv', 'w', newline='') as f_out:

                reader = csv.reader(f_in, delimiter=';')
                writer = csv.writer(f_out, delimiter=';')

                # Percorrer cada linha do arquivo de entrada
                for row in reader:
                    # Substituir todos os pontos por vírgulas na coluna desejada
                    row[4] = row[4].replace('.', ',')
                    row[6] = row[6].replace('.', '')
                    row[7] = row[7].replace('.', '')
                    row[8] = row[8].replace('.', '')
                    row[9] = row[9].replace('.', '')
                    row[10] = row[10].replace('.', '')
                    row[11] = row[11].replace('.', '')
                    row[12] = valorFaturaEntry
                    # Escrever a linha modificada no arquivo de saída
                    writer.writerow(row)

            # Exibir uma mensagem de conclusão
            messagebox.showinfo(
                'Concluído', 'O arquivo foi processado com sucesso.')
            root.destroy()

        root = tk.Tk()
        root.resizable(width=False, height=False)
        label = tk.Label(root, text="Fatura:")
        label.pack()
        entry = tk.Entry(root)
        entry.pack()
        # Criar botão para obter o valor
        btn_obter_valor = tk.Button(
            root, text="Enviar dados", command=obter_valor)
        btn_obter_valor.pack()

        # Executar o loop principal da janela
        root.mainloop()


def processar_arquivo_ALTICE():

    # Abrir a caixa de diálogo de seleção de arquivo
    filename = filedialog.askopenfilename(
        initialdir='/', title='Selecione o arquivo', filetypes=[('Arquivos do Excel', '*.csv')])
    if (filename == ''):
        messagebox.showinfo('Erro Sem Ficheiro',
                            'Nenhum arquivo foi selecionado.')
    else:
        def obter_valor():
            valorMES = entry.get_date()
            nome_arquivo, extensao = os.path.splitext(
                os.path.basename(filename))
            with open(filename, newline='') as f_in, open('C:\\importacao\\' + nome_arquivo + '.csv', 'w', newline='') as f_out:

                # Criar um leitor CSV e um gravador CSV
                leitor = csv.reader(f_in, delimiter=';')
                gravador = csv.writer(f_out, delimiter=';')

                rows = list(leitor)[6:-3]
                gravador.writerows(rows)
            with open('C:\\importacao\\' + nome_arquivo + '.csv', 'rb') as f:
                result = chardet.detect(f.read())
                encoding = result['encoding']

            df = pd.read_csv('C:\\importacao\\' + nome_arquivo +
                             '.csv', encoding=encoding, delimiter=";")

            df.to_excel('C:\\importacao\\' + nome_arquivo +
                        '.xlsx', index=False)

            os.remove('C:\\importacao\\' + nome_arquivo + '.csv')
            # Exibir uma mensagem de conclusão
            df = pd.read_excel('C:\\importacao\\' + nome_arquivo +
                               '.xlsx')

            teste = df.loc[0, "Plano de Preços"]

            num_rows = df.shape[0]
            num_rows2 = float(num_rows + 1)
            numero = float(teste.replace(',', '.'))
            media = numero/num_rows2

            df["VALOR MENSALIDADE"] = media
            valor = df["Valor (s/IVA)"].str.replace(',', '.').astype(float)
            df['VALOR DO CARTAO'] = media + valor

            df['DATA FATURA'] = valorMES

            df = df.iloc[1:]
            df.to_excel('C:\\importacao\\' + nome_arquivo +
                        '.xlsx', index=False)
            messagebox.showinfo(
                'Concluído', 'O arquivo foi processado com sucesso.')
            root.destroy()

    # Criar janela principal
    root = tk.Tk()
    root.resizable(width=False, height=False)

    # Criar rótulos
    label1 = tk.Label(root, text="Data:")
    label1.pack()
    entry = DateEntry(root, selectmode="day", date_pattern='yyyy-mm-dd')
    entry.pack()

    # Criar botão para obter o valor
    btn_obter_valor = tk.Button(
        root, text="Enviar dados", command=obter_valor)
    btn_obter_valor.pack()

    # Executar o loop principal da janela
    root.mainloop()


def processar_arquivo_VIAVERDE():

    # Abrir a caixa de diálogo de seleção de arquivo
    filename = filedialog.askopenfilename(
        initialdir='/', title='Selecione o arquivo', filetypes=[('Arquivos do Excel', '*.csv')])
    if (filename == ''):
        messagebox.showinfo('Erro Sem Ficheiro',
                            'Nenhum arquivo foi selecionado.')
    else:
        nome_arquivo, extensao = os.path.splitext(os.path.basename(filename))
        with open(filename, newline='') as f_in, open('C:\\importacao\\' + nome_arquivo + '.csv', 'w', newline='') as f_out:

            # Criar um leitor CSV e um gravador CSV
            leitor = csv.reader(f_in)
            gravador = csv.writer(f_out)

            # Iterar sobre as linhas do arquivo original e escrever no arquivo de saída
            for i, linha in enumerate(leitor):
                if i >= 7:  # Excluir as primeiras sete linhas
                    gravador.writerow(linha)
            # Ler o arquivo CSV
        with open('C:\\importacao\\' + nome_arquivo + '.csv', 'rb') as f:
            result = chardet.detect(f.read())
            encoding = result['encoding']

        df = pd.read_csv('C:\\importacao\\' + nome_arquivo +
                         '.csv', encoding=encoding, delimiter=";")

        dados1 = df.loc[(df['OPERADOR'] == 'B2') | (df['OPERADOR'] == 'E1') | (df['OPERADOR'] == 'TM') |
                        (df['OPERADOR'] == 'P3') | (df['OPERADOR'] == 'VI') | (df['OPERADOR'] == 'B1') |
                        (df['OPERADOR'] == 'P1') | (df['OPERADOR'] == 'O1') | (df['OPERADOR'] == 'VD') |
                        (df['OPERADOR'] == 'N1') | (df['OPERADOR'] == 'I1') | (df['OPERADOR'] == 'IF') |
                        (df['OPERADOR'] == 'E2') | (df['OPERADOR'] == 'BP') | (df['OPERADOR'] == 'P2') |
                        (df['OPERADOR'] == 'L1') | (df['OPERADOR'].str.contains('I. de Portugal'))]

        dados1.loc[:, 'OPERADOR'] = 'Infraestruturas de Portugal'
        dados1.loc[:, 'DATA ENTRADA'] = pd.to_datetime(
            dados1['DATA ENTRADA'], format='%d/%m/%Y')
        dados1.loc[:, 'DATA SAÍDA'] = pd.to_datetime(
            dados1['DATA SAÍDA'], format='%d/%m/%Y')
        dados1.loc[:, 'DATA PAGAMENTO'] = pd.to_datetime(
            dados1['DATA PAGAMENTO'], format='%d/%m/%Y')

        dados2 = df.loc[(df['OPERADOR'] == 'BR') | (
            df['OPERADOR'].str.contains('Brisa'))]

        dados2.loc[:, 'OPERADOR'] = 'Brisa Concessao Rodoviaria, S.'
        dados2.loc[:, 'DATA ENTRADA'] = pd.to_datetime(
            dados2['DATA ENTRADA'], format='%d/%m/%Y')
        dados2.loc[:, 'DATA SAÍDA'] = pd.to_datetime(
            dados2['DATA SAÍDA'], format='%d/%m/%Y')
        dados2.loc[:, 'DATA PAGAMENTO'] = pd.to_datetime(
            dados2['DATA PAGAMENTO'], format='%d/%m/%Y')

        dados3 = df.loc[(df['OPERADOR'] == 'S1') | (
            df['OPERADOR'].str.contains('Scutvias'))]

        dados3.loc[:, 'OPERADOR'] = 'Scutvias - Autoestradas da Beira, S.A.'
        dados3.loc[:, 'DATA ENTRADA'] = pd.to_datetime(
            dados3['DATA ENTRADA'], format='%d/%m/%Y')
        dados3.loc[:, 'DATA SAÍDA'] = pd.to_datetime(
            dados3['DATA SAÍDA'], format='%d/%m/%Y')
        dados3.loc[:, 'DATA PAGAMENTO'] = pd.to_datetime(
            dados3['DATA PAGAMENTO'], format='%d/%m/%Y')

        dados4 = df.loc[(df['OPERADOR'].str.contains('BRAGAPARQUES'))]

        dados4.loc[:, 'OPERADOR'] = 'Bragaparques, S.A.'
        dados4.loc[:, 'DATA ENTRADA'] = pd.to_datetime(
            dados4['DATA ENTRADA'], format='%d/%m/%Y')
        dados4.loc[:, 'DATA SAÍDA'] = pd.to_datetime(
            dados4['DATA SAÍDA'], format='%d/%m/%Y')
        dados4.loc[:, 'DATA PAGAMENTO'] = pd.to_datetime(
            dados4['DATA PAGAMENTO'], format='%d/%m/%Y')

        dados5 = df.loc[(df['OPERADOR'] == 'AA') | (
                        df['OPERADOR'].str.contains('AUTOESTRADAS DO ATLÂNTICO'))]

        dados5.loc[:, 'OPERADOR'] = 'AUTOESTRADAS DO ATLANTICO'
        dados5.loc[:, 'DATA ENTRADA'] = pd.to_datetime(
            dados5['DATA ENTRADA'], format='%d/%m/%Y')
        dados5.loc[:, 'DATA SAÍDA'] = pd.to_datetime(
            dados5['DATA SAÍDA'], format='%d/%m/%Y')
        dados5.loc[:, 'DATA PAGAMENTO'] = pd.to_datetime(
            dados5['DATA PAGAMENTO'], format='%d/%m/%Y')

        dados6 = df.loc[(df['OPERADOR'] == 'DL') | (
            df['OPERADOR'].str.contains('AEDL'))]

        dados6.loc[:, 'OPERADOR'] = 'Aedl - Estradas de Douro Litoral S.A.'
        dados6.loc[:, 'DATA ENTRADA'] = pd.to_datetime(
            dados6['DATA ENTRADA'], format='%d/%m/%Y')
        dados6.loc[:, 'DATA SAÍDA'] = pd.to_datetime(
            dados6['DATA SAÍDA'], format='%d/%m/%Y')
        dados6.loc[:, 'DATA PAGAMENTO'] = pd.to_datetime(
            dados6['DATA PAGAMENTO'], format='%d/%m/%Y')

        dados7 = df.loc[(df['OPERADOR'].str.contains('VV')) | (
            df['OPERADOR'].str.contains('VIA VERDE'))]

        dados7.loc[:, 'OPERADOR'] = 'Via Verde Portugal, S.A.'
        dados7.loc[:, 'DATA ENTRADA'] = pd.to_datetime(
            dados7['DATA ENTRADA'], format='%d/%m/%Y')
        dados7.loc[:, 'DATA SAÍDA'] = pd.to_datetime(
            dados7['DATA SAÍDA'], format='%d/%m/%Y')
        dados7.loc[:, 'DATA PAGAMENTO'] = pd.to_datetime(
            dados7['DATA PAGAMENTO'], format='%d/%m/%Y')

        dados8 = df.loc[(df['OPERADOR'] == 'VE')]

        dados8.loc[:, 'OPERADOR'] = 'Via Verde Pot.(Ve) Espanha'
        dados8.loc[:, 'DATA ENTRADA'] = pd.to_datetime(
            dados8['DATA ENTRADA'], format='%d/%m/%Y')
        dados8.loc[:, 'DATA SAÍDA'] = pd.to_datetime(
            dados8['DATA SAÍDA'], format='%d/%m/%Y')
        dados8.loc[:, 'DATA PAGAMENTO'] = pd.to_datetime(
            dados8['DATA PAGAMENTO'], format='%d/%m/%Y')

        dados9 = df.loc[(df['OPERADOR'] == 'LS') | (
            df['OPERADOR'].str.contains('Lusoponte'))]

        dados9.loc[:, 'OPERADOR'] = 'Lusoponte Concessionario para Trave.Tejo'
        dados9.loc[:, 'DATA ENTRADA'] = pd.to_datetime(
            dados9['DATA ENTRADA'], format='%d/%m/%Y')
        dados9.loc[:, 'DATA SAÍDA'] = pd.to_datetime(
            dados9['DATA SAÍDA'], format='%d/%m/%Y')
        dados9.loc[:, 'DATA PAGAMENTO'] = pd.to_datetime(
            dados9['DATA PAGAMENTO'], format='%d/%m/%Y')

        dados10 = df.loc[(df['OPERADOR'] == 'BL') | (
            df['OPERADOR'].str.contains('BRISAL'))]

        dados10.loc[:, 'OPERADOR'] = 'Brisal - Auto Estrada Do Litoral'
        dados10.loc[:, 'DATA ENTRADA'] = pd.to_datetime(
            dados10['DATA ENTRADA'], format='%d/%m/%Y')
        dados10.loc[:, 'DATA SAÍDA'] = pd.to_datetime(
            dados10['DATA SAÍDA'], format='%d/%m/%Y')
        dados10.loc[:, 'DATA PAGAMENTO'] = pd.to_datetime(
            dados10['DATA PAGAMENTO'], format='%d/%m/%Y')

        # Salvar os dados filtrados em arquivos CSV separados
        if len(dados1) != 0:
            dados1.to_excel('C:\\importacao\\' + nome_arquivo +
                            '_I_Portugal.xlsx', index=False)
        if len(dados2) != 0:
            dados2.to_excel('C:\\importacao\\' + nome_arquivo +
                            '_Brisa.xlsx', index=False)
        if len(dados3) != 0:
            dados3.to_excel('C:\\importacao\\' + nome_arquivo +
                            '_Scutvias.xlsx', index=False)
        if len(dados4) != 0:
            dados4.to_excel('C:\\importacao\\' + nome_arquivo +
                            '_BRAGAPARQUES.xlsx', index=False)
        if len(dados5) != 0:
            dados5.to_excel('C:\\importacao\\' + nome_arquivo +
                            '_AUTOESTRADAS_ATLANTICO.xlsx', index=False)
        if len(dados6) != 0:
            dados6.to_excel('C:\\importacao\\' + nome_arquivo +
                            '_AEDL.xlsx', index=False)
        if len(dados7) != 0:
            dados7.to_excel('C:\\importacao\\' + nome_arquivo +
                            '_VIA_VERDE_PORTUGAL.xlsx', index=False)
        if len(dados8) != 0:
            dados8.to_excel('C:\\importacao\\' + nome_arquivo +
                            '_VIA_VERDE_ESPANHA.xlsx', index=False)
        if len(dados9) != 0:
            dados9.to_excel('C:\\importacao\\' + nome_arquivo +
                            '_LUSOPONTE.xlsx', index=False)
        if len(dados10) != 0:
            dados10.to_excel('C:\\importacao\\' + nome_arquivo +
                             '_BRISAL.xlsx', index=False)

    os.remove('C:\\importacao\\' + nome_arquivo + '.csv')

    # Exibir uma mensagem de conclusão
    messagebox.showinfo(
        'Concluído', 'O arquivo foi processado com sucesso.')


def processar_arquivo_STARRESSA_ESPANHA_GASOLEO():
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

        duplicate_rows = df.duplicated()

        # Check if there are any duplicates
        if duplicate_rows.any():
            # Get the duplicate rows
            duplicate_data = df[duplicate_rows]
            print(duplicate_data)
            # Remove the duplicates
            df.drop_duplicates(inplace=True)

        # Calcular os Valores das colunas
        df['Valor Faturado'] = (
            df['Montante Operação'] - df['Montante desconto'])
        df['Valor Faturado S/ Iva '] = (df['Valor Faturado'] /
                                        ((df['% IMPOSTO']/100) + 1))

        # Exportar o DataFrame para um arquivo XLSX com as colunas selecionadas
        df.to_excel('C:\\importacao\\' + nome_arquivo + '.xlsx', index=False)
        # Exibir uma mensagem de conclusão
        messagebox.showinfo(
            'Concluído', 'O arquivo foi processado com sucesso.')

        # Exportar o DataFrame para um arquivo XLSX com as colunas selecionadas
        df.to_excel('C:\\importacao\\' + nome_arquivo + '.xlsx', index=False)
        # Exibir uma mensagem de conclusão
        messagebox.showinfo(
            'Concluído', 'O arquivo foi processado com sucesso.')


def processar_arquivo_STARRESSA_PORTUGAL_GASOLEO():

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
        duplicate_rows = df.duplicated()

        # Check if there are any duplicates
        if duplicate_rows.any():
            # Get the duplicate rows
            duplicate_data = df[duplicate_rows]
            print(duplicate_data)
            # Remove the duplicates
            df.drop_duplicates(inplace=True)

        # Calcular os Valores das colunas
        df['Total c / IVA'] = (df['Litros / Unidades']
                               * df['Valor do desconto'])

        # Exportar o DataFrame para um arquivo XLSX com as colunas selecionadas
        df.to_excel('C:\\importacao\\' + nome_arquivo + '.xlsx', index=False)
        # Exibir uma mensagem de conclusão
        messagebox.showinfo(
            'Concluído', 'O arquivo foi processado com sucesso.')


def processar_arquivo_ILIDIO_MOTA():
    # Abrir a caixa de diálogo de seleção de arquivo
    filename = filedialog.askopenfilename(
        initialdir='/', title='Selecione o arquivo', filetypes=[('Arquivos do Excel', '*.xlsx')])
    if (filename == ''):
        messagebox.showinfo('Erro Sem Ficheiro',
                            'Nenhum arquivo foi selecionado.')
    else:

        # Carregar o arquivo Excel em um DataFrame
        df = openpyxl.load_workbook(filename)
        # Remove o caminho e a extensao do nome do ficheiro
        nome_arquivo, extensao = os.path.splitext(os.path.basename(filename))

        # Seleciona a planilha
        planilha = df.active

        planilha.delete_cols(1)

        # Remover primeiras 5 linhas
        for i in range(1, 4):
            planilha.delete_rows(1)

        df.save(
            'C:\\importacao\\' + nome_arquivo + '.xlsx')

        df = pd.read_excel('C:\\importacao\\' + nome_arquivo + '.xlsx')
        linhas_para_remover = df[df["Design"] == "OFERTA CAFÉ"].index

        # Remove as linhas do DataFrame
        df.drop(linhas_para_remover, inplace=True)

        df.to_excel(
            'C:\\importacao\\' + nome_arquivo + '.xlsx', index=False)
        # Exibir uma mensagem de conclusão
        messagebox.showinfo(
            'Concluído', 'O arquivo foi processado com sucesso.')


def processar_arquivo_STARRESSA_FRANCA_PORTAGENS():
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

        # cabecalho = list(df)
        # Calcular os Valores das colunas
        def calcular_valor(row):
            if row['% IMPOSTO'] > 0:
                if row['Conceito'] == 'AUTOROUTES A PEAGE':
                    count = df.loc[df['Conceito'].isin(
                        ['AUTOROUTES A PEAGE'])].shape[0]

                    return (row['Montante Operação'] + (df.loc[df['Conceito'] == 'GEST. DES SERVICES AUTOROUTES', 'Montante Operação'].iloc[0] / count)) / (1 + (row['% IMPOSTO'] / 100))

                elif row['Conceito'] == 'GEST. DES SERVICES AUTOROUTES':
                    return (row['Montante Operação'])
                else:
                    return (row['Montante Operação'] / (1 + (row['% IMPOSTO'] / 100)))

            else:
                return row['Montante Operação']

        df['Comissões'] = df.apply(calcular_valor, axis=1)

        # Exportar o DataFrame para um arquivo XLSX com as colunas selecionadas
        df.to_excel('C:\\importacao\\' + nome_arquivo + '.xlsx', index=False)
        # Exibir uma mensagem de conclusão
        messagebox.showinfo(
            'Concluído', 'O arquivo foi processado com sucesso.')


def processar_arquivo_STARRESSA_ITALIA_PORTAGENS():
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

        # Calcular os Valores das colunas
        def calcular_valor(row):
            if row['% IMPOSTO'] > 0:
                if row['Conceito'] == 'PEDAGGI AUTOSTRADALI':
                    count = df.loc[df['Conceito'].isin(
                        ['PEDAGGI AUTOSTRADALI'])].shape[0]

                    return (row['Montante Operação'] + (df.loc[df['Conceito'] == 'GEST. SERV. PEAJES CONSORZIO', 'Montante Operação'].iloc[0] / count)) / (1 + (row['% IMPOSTO'] / 100))

                elif row['Conceito'] == 'GEST. SERV. PEAJES CONSORZIO':
                    return (row['Montante Operação'])
                else:
                    return (row['Montante Operação'] / (1 + (row['% IMPOSTO'] / 100)))

            else:
                return row['Montante Operação']

        df['Comissões'] = df.apply(calcular_valor, axis=1)

        # Exportar o DataFrame para um arquivo XLSX com as colunas selecionadas
        df.to_excel('C:\\importacao\\' + nome_arquivo + '.xlsx', index=False)
        # Exibir uma mensagem de conclusão
        messagebox.showinfo(
            'Concluído', 'O arquivo foi processado com sucesso.')


def processar_arquivo_STARRESSA_SUICA_PORTAGENS():
    # Abrir a caixa de diálogo de seleção de arquivo
    filename = filedialog.askopenfilename(
        initialdir='/', title='Selecione o arquivo', filetypes=[('Arquivos do Excel', '*.xlsx')])
    if (filename == ''):
        messagebox.showinfo('Erro Sem Ficheiro',
                            'Nenhum arquivo foi selecionado.')
    else:
        def obter_valor():
            valorFaturaEntry = entry.get()
            # Carregar o arquivo Excel em um DataFrame
            df = pd.read_excel(filename)
            # Remove o caminho e a extensao do nome do ficheiro
            nome_arquivo, extensao = os.path.splitext(
                os.path.basename(filename))

            df["MOEDA"] = "EURO"
            df["Montante Operação"] = (
                float(valorFaturaEntry)/100) * df["Montante Operação"]

            # Calcular os Valores das colunas
            def calcular_valor(row):
                if row['% IMPOSTO'] > 0:
                    if row['Conceito'] == 'AUTOROUTES A PEAGE':
                        count = df.loc[df['Conceito'].isin(
                            ['AUTOROUTES A PEAGE'])].shape[0]

                        return (row['Montante Operação'] + (df.loc[df['Conceito'] == 'GEST. DES SERVICES AUTOROUTES', 'Montante Operação'].iloc[0] / count)) / (1 + (row['% IMPOSTO'] / 100))

                    elif row['Conceito'] == 'GEST. DES SERVICES AUTOROUTES':
                        return (row['Montante Operação'])
                    else:
                        return (row['Montante Operação'] / (1 + (row['% IMPOSTO'] / 100)))

                else:
                    return row['Montante Operação']

            df['Comissões'] = df.apply(calcular_valor, axis=1)

            # Exportar o DataFrame para um arquivo XLSX com as colunas selecionadas
            df.to_excel('C:\\importacao\\' +
                        nome_arquivo + '.xlsx', index=False)
            # Exibir uma mensagem de conclusão
            messagebox.showinfo(
                'Concluído', 'O arquivo foi processado com sucesso.')
            root.destroy()

        # Criar janela principal
        root = tk.Tk()
        root.resizable(width=False, height=False)

        label = tk.Label(root, text="100 Francos - EUROS")
        label.pack()
        entry = tk.Entry(root)
        entry.pack()

        # Criar botão para obter o valor
        btn_obter_valor = tk.Button(
            root, text="Enviar dados", command=obter_valor)
        btn_obter_valor.pack()
        # Executar o loop principal da janela
        root.mainloop()


def processar_arquivo_IDS():

    filename = filedialog.askopenfilename(
        initialdir='/', title='Selecione o arquivo', filetypes=[('Arquivos do Excel', '*.xls')])
    if (filename == ''):
        messagebox.showinfo('Erro Sem Ficheiro',
                            'Nenhum arquivo foi selecionado.')
    else:
        def obter_valor():
            valorFaturaEntry = entry.get()
            # Carregar o arquivo Excel em um DataFrame
            nome_arquivo, extensao = os.path.splitext(
                os.path.basename(filename))
            df = pd.read_excel(filename)

            # Salvar como XLSX
            df.to_excel('C:\\importacao\\' +
                        nome_arquivo + '.xlsx', index=False)

            df = pd.read_excel('C:\\importacao\\' + nome_arquivo + '.xlsx')

            dados1 = df['TRS_DATE']
            df['TRS_DATE'] = pd.to_datetime(
                dados1, format='%Y-%m-%d').dt.date
            dados2 = df['INVO_DATE']
            df['INVO_DATE'] = pd.to_datetime(
                dados2, format='%Y-%m-%d').dt.date
            df["FATURA"] = valorFaturaEntry

            df.to_excel('C:\\importacao\\' +
                        nome_arquivo + '.xlsx', index=False)

            messagebox.showinfo(
                'Concluído', 'O arquivo foi processado com sucesso.')
            root.destroy()

        # Criar janela principal
        root = tk.Tk()

        # Criar widget Entry para entrada de texto
        label1 = tk.Label(root, text="Fatura:")
        label1.pack()
        entry = tk.Entry(root)
        entry.pack()
        # Criar botão para obter o valor
        btn_obter_valor = tk.Button(
            root, text="Enviar Valores", command=obter_valor)
        btn_obter_valor.pack()

        # Executar o loop principal da janela
        root.mainloop()


def processar_arquivo_MONTEPIO_RENTING():

    # Abrir a caixa de diálogo de seleção de arquivo
    filename = filedialog.askopenfilename(
        initialdir='/', title='Selecione o arquivo', filetypes=[('Arquivos do Excel', '*.xls')])
    if (filename == ''):
        messagebox.showinfo('Erro Sem Ficheiro',
                            'Nenhum arquivo foi selecionado.')
    else:
        def obter_valor():
            valorFaturaEntry = entry.get()
            valorFaturaEntry2 = entry2.get()

            # Ler o arquivo Excel
            nome_arquivo, extensao = os.path.splitext(
                os.path.basename(filename))
            df = pd.read_excel(filename)

            # Salvar como XLSX
            df.to_excel('C:\\importacao\\' +
                        nome_arquivo + '.xlsx', index=False)

            # Carregando o arquivo XLSX
            workbook = openpyxl.load_workbook(
                'C:\\importacao\\' + nome_arquivo + '.xlsx')

            worksheet = workbook['Sheet1']

            target_cell = None
            for row in worksheet.iter_rows():
                for cell in row:
                    if cell.value == "Matrícula":
                        target_cell = cell
                        break
                if target_cell:
                    break

            if target_cell:
                # Obter os dados abaixo da célula que contém o valor "Matrícula"
                data = []
                none_count = 0

                for row in worksheet.iter_rows(min_row=target_cell.row + 1, min_col=worksheet.min_column, max_col=worksheet.max_column):
                    row_data = [cell.value for cell in row]

                    none_count = 0

                    for value in row_data:
                        if value == 'Total':
                            none_count = 1
                            break
                        dados = [row_data[0], row_data[1], '', row_data[3],
                                 row_data[5], 'Aluguer', valorFaturaEntry2, valorFaturaEntry, '', '', '', '', '', '']

                    for value in row_data:
                        if value == 'Total':
                            none_count = 1
                            break
                        dados2 = [row_data[0], row_data[1], '', row_data[3], row_data[7],
                                  'Serviço de Gestão', valorFaturaEntry2, valorFaturaEntry, '', '', '', '', '', '']

                    for value in row_data:
                        if value == 'Total':
                            none_count = 1
                            break
                        dados3 = [row_data[0], row_data[1], '', row_data[3], row_data[8],
                                  'Contrato de Serviço', valorFaturaEntry2, valorFaturaEntry, '', '', '', '', '', '']

                    if none_count == 0:
                        data.append(dados)
                        data.append(dados2)
                        data.append(dados3)

                if data:
                    # Criar um DataFrame pandas com os dados
                    df = pd.DataFrame(
                        data, columns=[cell.value for cell in worksheet[1]])

                    # Escrever o DataFrame de volta no arquivo Excel
                    df.to_excel('C:\\importacao\\' + nome_arquivo +
                                '.xlsx', index=False)

            messagebox.showinfo(
                'Concluído', 'O arquivo foi processado com sucesso.')

            root.destroy()

        # Criar janela principal
        root = tk.Tk()

        # Criar widget Entry para entrada de texto
        label1 = tk.Label(root, text="Valor da Fatura:")
        label1.pack()
        entry = tk.Entry(root)
        entry.pack()
        label2 = tk.Label(root, text="Valor IVA:")
        label2.pack()
        entry2 = tk.Entry(root)
        entry2.pack()
        # Criar botão para obter o valor
        btn_obter_valor = tk.Button(
            root, text="Enviar Valores", command=obter_valor)
        btn_obter_valor.pack()

        # Executar o loop principal da janela
        root.mainloop()


def processar_arquivo_VALCARCE():

    # Abrir a caixa de diálogo de seleção de arquivo
    filename = filedialog.askopenfilename(
        initialdir='/', title='Selecione o arquivo', filetypes=[('Arquivos do Excel', '*.txt')])
    if (filename == ''):
        messagebox.showinfo('Erro Sem Ficheiro',
                            'Nenhum arquivo foi selecionado.')
    else:
        def obter_valor():
            valorFaturaEntry = entry.get_date()
            valorFaturaEntry2 = entry2.get()
            # Carregar o arquivo Excel em um DataFrame
            df = pd.read_csv(filename, delimiter='\t')
            # Remove o caminho e a extensao do nome do ficheiro
            nome_arquivo, extensao = os.path.splitext(
                os.path.basename(filename))

            # Exportar o DataFrame para um arquivo XLSX com as colunas selecionadas
            df.to_excel('C:\\importacao\\' +
                        nome_arquivo + '.xlsx', index=False)

            ficheiro = 'C:\\importacao\\' + nome_arquivo + '.xlsx'
            # Carregando o arquivo XLSX
            df = pd.read_excel(ficheiro)
            cabecalho = list(df)
            #  Altera as "," para "." para converter em float
            Precio = [s.replace(",", ".") for s in df['Pre.Clien']]
            Dto = [s.replace(",", ".") for s in df['Dto.Clien']]
            Cantidad = [s.replace(",", ".") for s in df['Cantidad']]

            df['Preco'] = df['Pre.Clien']
            df['Desconto'] = df['Dto.Clien']

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

            # Converte as colunas em float
            valores_Dto = []
            for valor in Dto:
                valor_float = float(valor)
                valores_Dto.append(valor_float)

            # Calcula os totais
            P_D = []
            for i in range(len(valores_Precio)):
                total = (valores_Precio[i] - valores_Dto[i])
                P_D.append(total)

            df['P-D'] = P_D

            # Calcula os totais
            PxQ = []
            for i in range(len(valores_Cantidad)):
                total = (P_D[i] * valores_Cantidad[i])
                PxQ.append(total)

            df['PxQ'] = PxQ

            # Seleciona a os elementos do Cabeçalho
            df['Mes de Faturacao'] = valorFaturaEntry
            df['Taxa de IVA'] = int(valorFaturaEntry2)

            cabecalho.append('Preco')
            cabecalho.append('Desconto')
            cabecalho.append('P-D')
            cabecalho.append('PxQ')
            cabecalho.append('Mes de Faturacao')
            cabecalho.append('Taxa de IVA')

            # Selecionar as colunas a serem exportadas
            colunas = cabecalho

            # Exportar o DataFrame para um arquivo XLSX com as colunas selecionadas
            df[colunas].to_excel(
                'C:\\importacao\\' + nome_arquivo + '.xlsx', index=False)

            # Exibir uma mensagem de conclusão
            messagebox.showinfo(
                'Concluído', 'O arquivo foi processado com sucesso.')

            root.destroy()

        # Criar janela principal
        root = tk.Tk()
        root.resizable(width=False, height=False)

        # Criar widget Entry para entrada de texto
        label1 = tk.Label(root, text="Mês de Faturação")
        label1.pack()
        entry = DateEntry(root, selectmode="day",
                          date_pattern='yyyy-mm-dd')
        entry.pack()
        label2 = tk.Label(root, text="Valor do IVA")
        label2.pack()
        entry2 = tk.Entry(root)
        entry2.pack()

        # Criar botão para obter o valor
        btn_obter_valor = tk.Button(
            root, text="Enviar dados", command=obter_valor)
        btn_obter_valor.pack()

        # Executar o loop principal da janela
        root.mainloop()


def processar_arquivo_SEGUROS():
    # Abrir a caixa de diálogo de seleção de arquivo
    filename = filedialog.askopenfilename(
        initialdir='/', title='Selecione o arquivo', filetypes=[('Arquivos do Excel', '*.xls')])
    if (filename == ''):
        messagebox.showinfo('Erro Sem Ficheiro',
                            'Nenhum arquivo foi selecionado.')
    else:
        # Ler o arquivo Excel
        nome_arquivo, extensao = os.path.splitext(
            os.path.basename(filename))
        df = pd.read_excel(filename)

        # Salvar como XLSX
        df.to_excel('C:\\importacao\\' + nome_arquivo + '.xlsx', index=False)

        # Carregar o arquivo XLSX
        df = pd.read_excel('C:\\importacao\\' + nome_arquivo + '.xlsx')

        df = df.dropna(subset=[df.columns[0]])
        primeiros_nove = df["Objecto Seguro"].str[:9]

        df["MATRICULA"] = primeiros_nove

        df.to_excel('C:\\importacao\\' + nome_arquivo +
                    '.xlsx', index=False)

    messagebox.showinfo(
        'Concluído', 'O arquivo foi processado com sucesso.')


def processar_arquivo_CONTRATOS_MAN():
    filename = filedialog.askopenfilename(
        initialdir='/', title='Selecione o arquivo', filetypes=[('Arquivos do Excel', '*.pdf')])
    if (filename == ''):
        messagebox.showinfo('Erro Sem Ficheiro',
                            'Nenhum arquivo foi selecionado.')
    else:
        # Caminho para o arquivo PDF
        pdf_path = filename

        # Extrair as tabelas do PDF usando o camelot-py
        nome_arquivo, extensao = os.path.splitext(os.path.basename(filename))
        tables = camelot.read_pdf(pdf_path, flavor="stream", pages="all")

        if tables:
            dfs = []
            for table in tables:
                df = table.df
                dfs.append(df)

            # Concatenar todas as tabelas em um único DataFrame
            final_df = pd.concat(dfs)

            # Salvar o DataFrame em um arquivo XLSX
            final_df.to_excel(
                'C:\\importacao\\' + nome_arquivo + '.xlsx', index=False)

        # Caminho para o arquivo Excel
        excel_path = 'C:\\importacao\\' + nome_arquivo + '.xlsx'

        # Carregar o arquivo Excel
        workbook = openpyxl.load_workbook(excel_path)

        # Selecionar a planilha desejada (substitua 'Sheet1' pelo nome da sua planilha)
        worksheet = workbook['Sheet1']

        # Encontrar a célula que contém o valor "Matrícula"
        target_cell = None
        for row in worksheet.iter_rows():
            for cell in row:
                if cell.value == "Matrícula":
                    target_cell = cell
                    break
            if target_cell:
                break

        if target_cell:
            # Obter a coluna correspondente à célula que contém o valor "Matrícula"
            col_index = target_cell.column

            # Obter os dados abaixo da célula que contém o valor "Matrícula"
            data = []
            none_count = 0

            for row in worksheet.iter_rows(min_row=target_cell.row + 1, min_col=worksheet.min_column, max_col=worksheet.max_column):
                row_data = [cell.value for cell in row]

                none_count = 0

                for value in row_data:
                    if value == None or value == 'N°CS' or value == 'N°chassis' or value == 'Tipo' or value == 'Nº Cliente':
                        none_count += 1

                if none_count < 4:
                    data.append(row_data)

            if data:
                # Criar um DataFrame pandas com os dados
                df = pd.DataFrame(
                    data, columns=[cell.value for cell in worksheet[1]])

                # Escrever o DataFrame de volta no arquivo Excel
                df.to_excel(excel_path, index=False)

            def valorIva():
                valorFaturaEntry = entry.get()
                valorFaturaEntry2 = entry2.get_date()
                valorFaturaEntry3 = entry3.get()
                # Carregar o arquivo Excel em um DataFrame
                df = pd.read_excel(excel_path)

                df['Data Fatura'] = valorFaturaEntry2

                df[6] = df[6].str.replace(',', '.').astype(float)
                df['IVA'] = int(valorFaturaEntry)
                df['Valor C/ Iva'] = (((df['IVA']/100) + 1) * df[6])
                df[4] = "MAN"
                df[5] = "Contrato Manutenção e Reparação MN Iva Taxa Normal"
                df["Nº Fatura"] = valorFaturaEntry3

                df.to_excel(excel_path, index=False)

                messagebox.showinfo(
                    'Concluído', 'O arquivo foi processado com sucesso.')
                root.destroy()

            # Criar janela principal
            root = tk.Tk()
            root.resizable(width=False, height=False)

            # Criar widget Entry para entrada de texto
            label1 = tk.Label(root, text="Valor do IVA")
            label1.pack()
            entry = tk.Entry(root)
            entry.pack()

            label3 = tk.Label(root, text="Nº da Fatura")
            label3.pack()
            entry3 = tk.Entry(root)
            entry3.pack()

            label2 = tk.Label(root, text="Dia Fatura")
            label2.pack()
            entry2 = DateEntry(root, selectmode="day",
                               date_pattern='yyyy-mm-dd')
            entry2.pack()

            # Criar botão para obter o valor
            btn_obter_valor = tk.Button(
                root, text="Enviar", command=valorIva)
            btn_obter_valor.pack()

            # Executar o loop principal da janela
            root.mainloop()


def processar_arquivo_AS24_FRANCA_PORTAGENS():
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

        # Exportar o DataFrame para um arquivo XLSX com as colunas selecionadas
        df.to_excel('C:\\importacao\\' + nome_arquivo + '.xlsx', index=False)
        # Exibir uma mensagem de conclusão
        messagebox.showinfo(
            'Concluído', 'O arquivo foi processado com sucesso.')


def processar_arquivo_AS24_FRANCA_COMBUSTIVEL():
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

        # Exportar o DataFrame para um arquivo XLSX com as colunas selecionadas
        df.to_excel('C:\\importacao\\' + nome_arquivo + '.xlsx', index=False)
        # Exibir uma mensagem de conclusão
        messagebox.showinfo(
            'Concluído', 'O arquivo foi processado com sucesso.')


def processar_arquivo_AS24_ESPANHA_PORTAGENS():
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

        # Exportar o DataFrame para um arquivo XLSX com as colunas selecionadas
        df.to_excel('C:\\importacao\\' + nome_arquivo + '.xlsx', index=False)
        # Exibir uma mensagem de conclusão
        messagebox.showinfo(
            'Concluído', 'O arquivo foi processado com sucesso.')


def processar_arquivo_BOMBA_PRÓPRIA_ABLUE_PARQUE():

    def obter_valor():
        valorMES = entry.get()
        valorANO = entry1.get()
        # Configuração da conexão com o SQL Server
        server = 'SBs2019-ISAAC\ABMN'
        database = 'aTrans'
        username = 'Bds'
        password = 'olivettiBDS1'

        # Criação da conexão com o SQL Server
        conn = pyodbc.connect(
            'DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD=' + password)

        # Parâmetros da stored procedure
        param1 = valorMES
        param2 = valorANO

        # Chamar a stored procedure com os parâmetros
        cursor = conn.cursor()
        cursor.execute(
            "{CALL sp_luzitic_dados_hecpoll (?, ?)}", (param1, param2))

        # Recuperar os resultados em uma lista de DataFrames
        tables = []

        while True:
            if cursor.description is not None:
                columns = [column[0] for column in cursor.description]
                rows = cursor.fetchall()
                df = pd.DataFrame.from_records(rows, columns=columns)
                tables.append(df)
            if not cursor.nextset():
                break

        # Fecha a conexão com o banco de dados
        conn.close()

        # Verifica se existem tabelas a serem salvas
        if len(tables) == 0:
            messagebox.showinfo(
                'ERRO', 'Nenhuma tabela encontrada para salvar.')
            return

        # Gera o arquivo Excel com as tabelas retornadas
        excel_file = 'C:\\importacao\\BOMBA_PRÓPRIA_ABLUE_PARQUE' + \
            valorMES+'_' + valorANO + '.xlsx'

        with pd.ExcelWriter(excel_file) as writer:
            for i, table in enumerate(tables):
                table.to_excel(
                    writer, sheet_name=f'Tabela {i+1}', index=False)

        df = pd.read_excel(excel_file)

        dados1 = df['Data']
        df['Data'] = pd.to_datetime(
            dados1, format='%Y-%m-%d').dt.date
        df["Quantidade"] = df["Quantidade"].replace('.', ',')
        df["Total_S_Iva"] = df["Total_S_Iva"].replace('.', ',')
        df["Total_C_Iva"] = df["Total_C_Iva"].replace('.', ',')

        # Margem do ADBLUE

        preco_Litro = (df["Total_C_Iva"] / df["Quantidade"])

        preco_Litro = (preco_Litro/((df["IVA"]/100)+1))

        pagaMais10 = (preco_Litro*1.10)
        pagaMais20 = (preco_Litro*1.20)

        pagaMais10 = (pagaMais10 * df['Quantidade'])
        pagaMais20 = (pagaMais20 * df['Quantidade'])

        df["Total_C_Iva"] = pagaMais20

        df.loc[(df['Cod_Artigo'] == 9) & (df['Empresa']
                                          == 'JPO'), 'Total_C_Iva'] = pagaMais10
        df.loc[(df['Cod_Artigo'] == 9) & (df['Empresa'] ==
                                          'ISAAC PEDROSO SRL'), 'Total_C_Iva'] = pagaMais10

        # Condições das Lavagens

        # Lavagem Completa Lonas
        df.loc[(df['Cod_Artigo'] == 2) & (
            df['Empresa'] == 'JPO'), 'Total_C_Iva'] = 30 * 1.23
        df.loc[(df['Cod_Artigo'] == 2) & (
            df['Empresa'] == 'JPO'), 'Total_S_Iva'] = 30

        df.loc[(df['Cod_Artigo'] == 2) & (df['Empresa']
                                          == 'F. FERNANDO LDA'), 'Total_C_Iva'] = 40 * 1.23
        df.loc[(df['Cod_Artigo'] == 2) & (df['Empresa']
                                          == 'F. FERNANDO LDA'), 'Total_S_Iva'] = 40

        df.loc[(df['Cod_Artigo'] == 2) & (df['Empresa']
                                          == 'TRANS JECHIU'), 'Total_C_Iva'] = 40 * 1.23
        df.loc[(df['Cod_Artigo'] == 2) & (df['Empresa']
                                          == 'TRANS JECHIU'), 'Total_S_Iva'] = 40

        # Lavagem Cisternas
        df.loc[(df['Cod_Artigo'] == 3) & (df['Empresa'] ==
                                          'ISAAC PEDROSO SRL'), 'Total_C_Iva'] = 40 * 1.23
        df.loc[(df['Cod_Artigo'] == 3) & (df['Empresa'] ==
                                          'ISAAC PEDROSO SRL'), 'Total_S_Iva'] = 40

        # Lavagem Autocarros
        df.loc[(df['Cod_Artigo'] == 6) & (
            df['Matricula'] == 'TRANSDEV'), 'Total_C_Iva'] = 30 * 1.23
        df.loc[(df['Cod_Artigo'] == 6) & (
            df['Matricula'] == 'TRANSDEV'), 'Total_S_Iva'] = 30

        # Lavagem Ligeiros
        df.loc[(df['Cod_Artigo'] == 8) & (df['Empresa'] ==
                                          'SEQUEIRA PEDROSO'), 'Total_C_Iva'] = 15 * 1.23
        df.loc[(df['Cod_Artigo'] == 8) & (df['Empresa'] ==
                                          'SEQUEIRA PEDROSO'), 'Total_S_Iva'] = 15
        df.loc[(df['Cod_Artigo'] == 8) & (
            df['Matricula'] == 'PANIPRADO'), 'Total_C_Iva'] = 20 * 1.23
        df.loc[(df['Cod_Artigo'] == 8) & (
            df['Matricula'] == 'PANIPRADO'), 'Total_S_Iva'] = 20

        df.to_excel(excel_file, index=False)

        messagebox.showinfo(
            'Concluído', 'O arquivo foi processado com sucesso.')
        root.destroy()

    # Criar janela principal
    root = tk.Tk()
    root.resizable(width=False, height=False)

    # Criar rótulos
    label1 = tk.Label(root, text="Mês: (em valor numérico)")
    label1.pack()
    entry = tk.Entry(root)
    entry.pack()

    label2 = tk.Label(root, text="Ano:")
    label2.pack()
    entry1 = tk.Entry(root)
    entry1.pack()

    # Criar botão para obter o valor
    btn_obter_valor = tk.Button(
        root, text="Enviar dados", command=obter_valor)
    btn_obter_valor.pack()

    # Executar o loop principal da janela
    root.mainloop()


def processar_arquivo_CTIB():
    filename = filedialog.askopenfilename(
        initialdir='/', title='Selecione o arquivo', filetypes=[('Arquivos do Excel', '*.pdf')])
    if (filename == ''):
        messagebox.showinfo('Erro Sem Ficheiro',
                            'Nenhum arquivo foi selecionado.')
    else:
        pdf_path = filename

        # Extrair as tabelas do PDF usando o camelot-py
        nome_arquivo, extensao = os.path.splitext(os.path.basename(filename))
        tables = camelot.read_pdf(pdf_path, flavor="stream", pages="all")

        if tables:
            dfs = []
            for table in tables:
                df = table.df
                dfs.append(df)

            # Concatenar todas as tabelas em um único DataFrame
            final_df = pd.concat(dfs)

            # Salvar o DataFrame em um arquivo XLSX
            final_df.to_excel(
                'C:\\importacao\\' + nome_arquivo + '.xlsx', index=False)

        # Caminho para o arquivo Excel
        excel_path = 'C:\\importacao\\' + nome_arquivo + '.xlsx'

        # Carregar o arquivo Excel
        workbook = openpyxl.load_workbook(excel_path)

        # Selecionar a planilha desejada (substitua 'Sheet1' pelo nome da sua planilha)
        worksheet = workbook['Sheet1']

        # Encontrar a célula que contém o valor "Matrícula"
        target_cell = None
        for row in worksheet.iter_rows():

            for cell in row:
                if cell.value == "Requisição:":
                    target_cell = cell
                    break
            if target_cell:
                break

        if target_cell:
            # Obter a coluna correspondente à célula que contém o valor "Matrícula"
            col_index = target_cell.column

            # Obter os dados abaixo da célula que contém o valor "Matrícula"
            data = []
            none_count = 0

            for row in worksheet.iter_rows(min_row=target_cell.row + 1, min_col=worksheet.min_column, max_col=worksheet.max_column):
                row_data = [cell.value for cell in row]

                none_count = 0

                for value in row_data:
                    if value == None or value == 'Motivo':
                        none_count += 1
                if none_count < 4:
                    data.append(row_data)

            if data:
                # Criar um DataFrame pandas com os dados
                df = pd.DataFrame(
                    data, columns=[cell.value for cell in worksheet[1]])

                df = df.dropna(how='all')
                df = df.reset_index(drop=True)
                # Escrever o DataFrame de volta no arquivo Excel
                df.to_excel(excel_path, index=False)

            def obter_valor():
                valorFaturaEntry = entry.get_date()
                valorFaturaEntry2 = entry2.get()
                valorFaturaEntry3 = entry3.get()
                # Read the Excel file
                df = pd.read_excel(excel_path)

                # Transpose the DataFrame to convert columns to rows
                # Obter o valor da célula
                valor4 = df.loc[0, 1]
                # Remove the header
                df = df.iloc[1:]

                # Remove the last row
                df = df.iloc[:-1]
                df = df.transpose()

                # Obter o valor da célula
                valor = df.loc[1, 7]
                valor2 = df.loc[2, 7]
                # Atribuir o valor da célula acima
                df.loc[0, 7] = valor
                df.loc[1, 7] = valor2

                # Atribuir um valor vazio à célula acima
                df.loc[2, 7] = pd.NA
                df[7] = df[7].str.replace('€', '')
                df['Motivo Inspecao'] = valor4
                df['IVA'] = int(valorFaturaEntry2)

                df = df.drop(0)
                valorIVA = (df['IVA']/100)+1
                valor5 = float(df.loc[1, 7].replace(',', '.'))
                df['Total S/IVA'] = valor5 / valorIVA
                df['Fatura'] = valorFaturaEntry3
                df['Data Fatura'] = valorFaturaEntry
                # Remove the last row
                df = df.iloc[:-3]
                # Save the transposed DataFrame to a new Excel file
                df.to_excel(excel_path, index=False)
                messagebox.showinfo(
                    'Concluído', 'O arquivo foi processado com sucesso.')
                root.destroy()

            # Criar janela principal
            root = tk.Tk()
            root.resizable(width=False, height=False)
            # Criar rótulos
            label1 = tk.Label(root, text="Data:")
            label1.pack()
            entry = DateEntry(root, selectmode="day",
                              date_pattern='yyyy-mm-dd')
            entry.pack()

            label3 = tk.Label(root, text="Valor do IVA:")
            label3.pack()
            entry2 = tk.Entry(root)
            entry2.pack()

            label2 = tk.Label(root, text="Fatura:")
            label2.pack()
            entry3 = tk.Entry(root)
            entry3.pack()

            # Criar botão para obter o valor
            btn_obter_valor = tk.Button(
                root, text="Enviar dados", command=obter_valor)
            btn_obter_valor.pack()

            # Executar o loop principal da janela
            root.mainloop()


def processar_arquivo_TRIMBLE():
    # Abrir a caixa de diálogo de seleção de arquivo
    def obter_valor():
        valorMES = entry.get_date()
        valorFATURA = entry3.get()

        valorData = str(valorMES)
        # Carregar o arquivo Excel em um DataFrame
        # Configuração da conexão com o SQL Server
        server = 'SBs2019-ISAAC\ABMN'
        database = 'trimble'
        username = 'Bds'
        password = 'olivettiBDS1'

        # Criação da conexão com o SQL Server
        conn = pyodbc.connect(
            'DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD=' + password)

        # Parâmetros da stored procedure
        # Chamar a stored procedure com os parâmetros
        cursor = conn.cursor()
        cursor.execute(
            "{CALL ImportacoesTrimble (?)}", (valorData))

        # Recuperar os resultados em uma lista de DataFrames
        tables = []

        while True:
            if cursor.description is not None:
                columns = [column[0] for column in cursor.description]
                rows = cursor.fetchall()
                df = pd.DataFrame.from_records(rows, columns=columns)
                tables.append(df)
            if not cursor.nextset():
                break

        # Fecha a conexão com o banco de dados
        conn.close()

        # Verifica se existem tabelas a serem salvas
        if len(tables) == 0:
            messagebox.showinfo(
                'ERRO', 'Nenhuma tabela encontrada para salvar.')
            return

        # Gera o arquivo Excel com as tabelas retornadas
        excel_file = 'C:\\importacao\\Trimble' + \
            valorData+'.xlsx'

        with pd.ExcelWriter(excel_file) as writer:
            for i, table in enumerate(tables):
                table.to_excel(
                    writer, sheet_name=f'Tabela {i+1}', index=False)
        df = pd.read_excel(excel_file)

        df['nrdoc'] = valorFATURA

        df['datafatura'] = valorMES.strftime('%Y-%m-%d')

        df.to_excel(excel_file, index=False)
        messagebox.showinfo(
            'Concluído', 'O arquivo foi processado com sucesso.')
        root.destroy()

    # Criar janela principal
    root = tk.Tk()
    root.resizable(width=False, height=False)

    # Criar rótulos
    label1 = tk.Label(root, text="Data:")
    label1.pack()
    entry = DateEntry(root, selectmode="day", date_pattern='yyyy-mm-dd')
    entry.pack()

    # label3 = tk.Label(root, text="Ano:")
    # label3.pack()
    # entry2 = tk.Entry(root)
    # entry2.pack()

    label2 = tk.Label(root, text="Valor Fatura:")
    label2.pack()
    entry3 = tk.Entry(root)
    entry3.pack()

    # Criar botão para obter o valor
    btn_obter_valor = tk.Button(
        root, text="Enviar dados", command=obter_valor)
    btn_obter_valor.pack()

    # Executar o loop principal da janela
    root.mainloop()

# FALTA RECEBER ORDEM DE CARGA


def processar_arquivo_WTRANSNET():

    filename = filedialog.askopenfilename(
        initialdir='/', title='Selecione o arquivo', filetypes=[('Arquivos do Excel', '*.pdf')])
    if (filename == ''):
        messagebox.showinfo('Erro Sem Ficheiro',
                            'Nenhum arquivo foi selecionado.')
    else:
        pdf_path = filename

        # Extrair as tabelas do PDF usando o camelot-py
        nome_arquivo, extensao = os.path.splitext(os.path.basename(filename))
        tables = camelot.read_pdf(pdf_path, flavor="stream", pages="all")

        if tables:
            dfs = []
            for table in tables:
                df = table.df
                dfs.append(df)

            # Concatenar todas as tabelas em um único DataFrame
            final_df = pd.concat(dfs)

            # Salvar o DataFrame em um arquivo XLSX
            final_df.to_excel(
                'C:\\importacao\\' + nome_arquivo + '.xlsx', index=False)

        # Caminho para o arquivo Excel
        excel_path = 'C:\\importacao\\' + nome_arquivo + '.xlsx'

        # Carregar o arquivo Excel
        workbook = openpyxl.load_workbook(excel_path)

        # Selecionar a planilha desejada (substitua 'Sheet1' pelo nome da sua planilha)
        worksheet = workbook['Sheet1']

        # Encontrar a célula que contém o valor "Matrícula"
        target_cell = None
        for row in worksheet.iter_rows():
            for cell in row:
                if cell.value == "Código":
                    target_cell = cell
                    break
            if target_cell:
                break

        if target_cell:
            # Obter a coluna correspondente à célula que contém o valor "Matrícula"
            col_index = target_cell.column

            # Obter os dados abaixo da célula que contém o valor "Matrícula"
            data = []
            none_count = 0

            for row in worksheet.iter_rows(min_row=target_cell.row + 1, min_col=worksheet.min_column, max_col=worksheet.max_column):
                row_data = [cell.value for cell in row]

                none_count = 0

                for value in row_data:
                    if value == None or value == 'N°CS' or value == 'N°chassis' or value == 'Tipo' or value == 'Nº Cliente':
                        none_count += 1

                if none_count < 4:
                    data.append(row_data)

            if data:
                # Criar um DataFrame pandas com os dados
                df = pd.DataFrame(
                    data, columns=[cell.value for cell in worksheet[1]])

                # Escrever o DataFrame de volta no arquivo Excel
                df.to_excel(excel_path, index=False)
            df = pd.read_excel(excel_path)
            df[7] = df[7].replace('.', ',')
            # Save the transposed DataFrame to a new Excel file
            df.to_excel(excel_path, index=False)
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
        def obter_valor():
            valorFATURA = entry3.get()
            # Carregar o arquivo Excel em um DataFrame
            df = pd.read_excel(filename, sheet_name="Details")
            # Remove o caminho e a extensao do nome do ficheiro
            nome_arquivo, extensao = os.path.splitext(
                os.path.basename(filename))

            df["ORDEM DE CARGA"] = "OC 132"
            df["NumeroFATURA"] = valorFATURA

            # Exportar o DataFrame para um arquivo XLSX com as colunas selecionadas
            df.to_excel('C:\\importacao\\' +
                        nome_arquivo + '.xlsx', index=False)
            # Exibir uma mensagem de conclusão
            messagebox.showinfo(
                'Concluído', 'O arquivo foi processado com sucesso.')
            root.destroy()

    # Criar janela principal
    root = tk.Tk()
    root.resizable(width=False, height=False)

    label2 = tk.Label(root, text="Numero da Fatura:")
    label2.pack()
    entry3 = tk.Entry(root)
    entry3.pack()

    # Criar botão para obter o valor
    btn_obter_valor = tk.Button(
        root, text="Enviar dados", command=obter_valor)
    btn_obter_valor.pack()


def selecionar_opcao(event):
    opcao = combo_box.get()

    if opcao == "STARRESSA - ESPANHA - GASÓLEO":
        processar_arquivo_STARRESSA_ESPANHA_GASOLEO()

    elif opcao == "STARRESSA - PORTUGAL - GASÓLEO":
        processar_arquivo_STARRESSA_PORTUGAL_GASOLEO()

    elif opcao == "STARRESSA - FRANÇA - PORTAGENS":
        processar_arquivo_STARRESSA_FRANCA_PORTAGENS()

    elif opcao == "STARRESSA - SUÍÇA - PORTAGENS":
        processar_arquivo_STARRESSA_SUICA_PORTAGENS()

    elif opcao == "STARRESSA - ITÁLIA - PORTAGENS":
        processar_arquivo_STARRESSA_ITALIA_PORTAGENS()

    elif opcao == "MONTEPIO - RENTING":
        processar_arquivo_MONTEPIO_RENTING()

    elif opcao == "VALCARCE":
        processar_arquivo_VALCARCE()

    elif opcao == "VIA VERDE":
        processar_arquivo_VIAVERDE()

    elif opcao == "ILIDIO MOTA":
        processar_arquivo_ILIDIO_MOTA()

    elif opcao == "NORPETROL":
        processar_arquivo_NORPETROL()

    elif opcao == "IDS":
        processar_arquivo_IDS()

    elif opcao == "SEGUROS":
        processar_arquivo_SEGUROS()

    elif opcao == "BOMBA PRÓPRIA - ABLUE PARQUE":
        processar_arquivo_BOMBA_PRÓPRIA_ABLUE_PARQUE()

    elif opcao == "TRIMBLE":
        processar_arquivo_TRIMBLE()

    elif opcao == "CONTRATOS DE MANUTENÇÃO - MAN":
        processar_arquivo_CONTRATOS_MAN()

    elif opcao == "AS24 - ESPANHA":
        processar_arquivo_AS24_ESPANHA_PORTAGENS()

    elif opcao == "AS24 - FRANÇA":
        processar_arquivo_AS24_FRANCA_COMBUSTIVEL()

    elif opcao == "Toll Collect":
        processar_arquivo_Toll_Collect()

    elif opcao == "ALTICE":
        processar_arquivo_ALTICE()

    elif opcao == "WTRANSNET":
        processar_arquivo_WTRANSNET()

    elif opcao == "VIALTIS":
        processar_arquivo_VIALTIS()

    else:
        processar_arquivo_PorFazer()


# Criar a janela principal
root = tk.Tk()

# desabilita a maximização e define o tamanho fixo da janela
root.resizable(width=False, height=False)
# Definir o tamanho da janela
root.geometry('500x300')

# Definir o título da janela
root.title('Carregamento de Arquivo Excel')

# Definir a cor de fundo da janela
root.configure(bg='#F0F0F0')

# Define o ícone da janela
root.iconbitmap("c:/Users/Marcos/Desktop/py/BOLA-LUZITIC.ico")

# Carrega a imagem
img = Image.open("c:/Users/Marcos/Desktop/py/luzitic.png")
img2 = Image.open("c:/Users/Marcos/Desktop/py/logo-isac.png")
# Cria o objeto ImageTk
img_tk = ImageTk.PhotoImage(img)
img_tk2 = ImageTk.PhotoImage(img2)

# Criar um botão "Selecionar arquivo"
# Ver Query para aqui
combo_box = tk.ttk.Combobox(
    root, state="readonly", values=["AS24 - ESPANHA",
                                    "AS24 - FRANÇA",
                                    "Toll Collect",
                                    "STARRESSA - ESPANHA - GASÓLEO",
                                    "STARRESSA - PORTUGAL - GASÓLEO",
                                    "STARRESSA - FRANÇA - PORTAGENS",
                                    "STARRESSA - ITÁLIA - PORTAGENS",
                                    "STARRESSA - SUÍÇA - PORTAGENS",
                                    "MONTEPIO - RENTING",
                                    "VALCARCE",
                                    "ALTICE",
                                    "VIA VERDE",
                                    "SEGUROS",
                                    "ILIDIO MOTA",
                                    "NORPETROL",
                                    "IDS",
                                    "CONTRATOS DE MANUTENÇÃO - MAN",
                                    "BOMBA PRÓPRIA - ABLUE PARQUE",
                                    "VIALTIS",
                                    "--------------------------------------------",
                                    "AS24 - PORTUGAL",
                                    "WTRANSNET",
                                    "TRIMBLE",
                                    "CONTRATOS DE MANUTENÇÃO  - IVECO",
                                    "CONTRATOS DE MANUTENÇÃO  - SCANIA",
                                    "BP",
                                    "REPSOL",
                                    "STARRESSA - ALEMANHA - PORTAGENS",
                                    "STARRESSA - BÉLGICA - PORTAGENS"])
combo_box.config(width=50)
combo_box.pack(pady=140)

combo_box.bind("<<ComboboxSelected>>", selecionar_opcao)


# Cria o rótulo e exibe a imagem
# Posiciona a imagem no canto superior direito
label = tk.Label(root, image=img_tk)
label.place(relx=1.0, rely=1.0, anchor="se")


# Cria o rótulo e exibe a imagem
# Posiciona a imagem no canto superior direito
label2 = tk.Label(root, image=img_tk2)
label2.place(x=0, y=10, anchor="nw")


texto = "Escolher Empresa para a Importação"
label3 = tk.Label(root, text=texto, font=("Arial", 14))
label3.place(relx=0.5, x=0, y=120, anchor="c")


label4 = tk.Label(root, text='Power By :', font=("Arial", 8))
label4.place(relx=0.75, x=0, y=280, anchor="c")


# Iniciar a janela
root.mainloop()
# Fim do Programa
