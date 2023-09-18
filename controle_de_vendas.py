#!/usr/bin/env python
# coding: utf-8
import pyodbc
from tkinter.messagebox import askyesno
from datetime import date
import pandas as pd
from babel.numbers import format_currency
import tkinter.filedialog
from tkinter import *
from tkinter import ttk
import tkinter as tk

janela = Tk()
arquivo = tkinter.filedialog.askopenfilename(title="Selecione o Arquivo que possui os códigos dos produtos")
janela.destroy()

dados_conexao = 'Driver={SQLite3 ODBC Driver};Server=localhost;Database=cavaleiro_sombrio.db'

conexao = pyodbc.connect(dados_conexao)
cursor = conexao.cursor()


def remover_item():
    cursor.execute('''
    DELETE FROM vendas_bar WHERE ID=(SELECT MAX(id) FROM vendas_bar)
    ''')
    cursor.commit()
    caixa_texto.delete("1.0", END)
    caixa_texto.insert("1.0", 'O último item adicionado foi removido.')


def zerar_janela():
    root = tk.Tk()

    canvas1 = tk.Canvas(root, width=300, height=300)
    canvas1.pack()
    msg_box = tk.messagebox.askquestion('Confirmação', 'Deseja deletar a tabela?',
                                        icon='warning')
    if msg_box == 'yes':
        cursor.execute('''
DELETE FROM vendas_bar 
''')
        cursor.commit()
        caixa_texto.delete("1.0", END)
        caixa_texto.insert("1.0", 'Tabela resetada.')
        root.destroy()
    else:
        tk.messagebox.showinfo(' ', 'Você voltara para a tela principal')
        root.destroy()

    root.mainloop()


def exporta_excel():
    hoje = date.today()

    codigo_df = pd.read_excel(arquivo)
    cursor.execute('SELECT * FROM vendas_bar')

    valores = cursor.fetchall()
    descricao = cursor.description

    colunas = [tupla[0] for tupla in descricao]
    tabela_clientes = pd.DataFrame.from_records(valores, columns=colunas)

    tabela_clientes['CODIGOPRODUTO'] = tabela_clientes['CODIGOPRODUTO'].astype(int)
    tabela_clientes['QUANTIDADE'] = tabela_clientes['QUANTIDADE'].astype(int)

    tabela_clientes = tabela_clientes.drop(columns=['ID'], axis=1)

    tabela_clientes = tabela_clientes.merge(codigo_df, on='CODIGOPRODUTO', how='left')

    tabela_clientes['VALOR FINAL'] = tabela_clientes['QUANTIDADE'] * tabela_clientes['VALOR']

    tabela_clientes['VALOR FINAL'] = tabela_clientes['VALOR FINAL'].apply(
        lambda x: format_currency(x, currency="BRL", locale="nl_NL"))
    tabela_clientes['VALOR'] = tabela_clientes['VALOR'].apply(
        lambda x: format_currency(x, currency="BRL", locale="nl_NL"))
    tabela_clientes.rename(columns={'VALOR': 'VALOR UNIT.'}, inplace=True)

    tabela_clientes = tabela_clientes[
        ['CODIGOPRODUTO', 'PRODUTO', 'QUANTIDADE', 'VALOR UNIT.', 'VALOR FINAL', 'FORMADEPAGAMENTO']]
    tabela_clientes.to_csv(r'{}-Vendas-Bar.csv'.format(hoje), sep=';', encoding='latin1', index=False)

    caixa_texto.delete("1.0", END)
    caixa_texto.insert("1.0", 'Tabela exportada.')
    cursor.commit()


def adicionar_produto():
    codigo_df = pd.read_excel(arquivo)
    codigo = codigo_produto.get()
    quantidade = quantidade_prod.get()
    cursor.execute('SELECT * FROM vendas_bar')

    if len(codigo) == 3:
        if not codigo.isnumeric():
            caixa_texto.delete("1.0", END)
            caixa_texto.insert("1.0", 'Código do produto inválido')
        elif not quantidade.isnumeric():
            caixa_texto.delete("1.0", END)
            caixa_texto.insert("1.0", '''
    Preencha a quantidade corretamente. 
    Somente com números.''')
        elif pagamento.get() == '':
            caixa_texto.delete("1.0", END)
            caixa_texto.insert("1.0", '''
    Selecione a forma de pagamento.''')
        else:

            valor_prod = codigo_df.loc[codigo_df['CODIGOPRODUTO'] == int(codigo_produto.get()), 'VALOR'].item()
            nome_produto = codigo_df.loc[codigo_df['CODIGOPRODUTO'] == int(codigo_produto.get()), 'PRODUTO'].item()
            valor_total = valor_prod * int(quantidade_prod.get())
            cursor.execute(f'''
            INSERT INTO vendas_bar (CODIGOPRODUTO, QUANTIDADE, FORMADEPAGAMENTO)
            VALUES
            ({codigo_produto.get()}, {quantidade_prod.get()}, "{pagamento.get()}")
            ''')
            cursor.commit()
            # deletar tudo da caixa de texto
            caixa_texto.delete("1.0", END)

            # escrever na caixa de texto
            caixa_texto.insert("1.0",
                               f'Produto {codigo_produto.get()} - {nome_produto}\n'
                               f'{quantidade_prod.get()} unidades no total de R${valor_total:,.2f}\n'
                               f'Forma de pagamento: {pagamento.get()}')
            codigo_produto.delete(0, END)
            quantidade_prod.delete(0, END)
    else:
        caixa_texto.delete("1.0", END)
        caixa_texto.insert("1.0", 'Código do produto inválido')


# Janela principal
window = Tk()

window.geometry("900x720")
window.configure(bg="#ffffff")
canvas = Canvas(
    window,
    bg="#ffffff",
    height=720,
    width=900,
    bd=0,
    highlightthickness=0,
    relief="ridge")
canvas.place(x=0, y=0)

background_img = PhotoImage(file=f"janela/templateCAV.png")
background = canvas.create_image(
    450, 360,
    image=background_img)

# Botão de adicionar tabela
img0 = PhotoImage(file=f"janela/B_add.png")
adicionar = Button(
    image=img0,
    borderwidth=0,
    highlightthickness=0,
    command=adicionar_produto,
    relief="flat")

adicionar.place(
    x=240, y=380,
    width=178,
    height=32)

# Botão de exportar tabela
img1 = PhotoImage(file=f"janela/B_EXP.png")
exportar = Button(
    image=img1,
    borderwidth=0,
    highlightthickness=0,
    command=exporta_excel,
    relief="flat")

exportar.place(
    x=440, y=380,
    width=178,
    height=32)

# Botão de zerar tabela
img2 = PhotoImage(file=f"janela/RESETAR.png")
zerar = Button(
    image=img2,
    borderwidth=0,
    highlightthickness=0,
    command=zerar_janela,
    relief="flat")

zerar.place(
    x=640, y=380,
    width=178,
    height=32)

# Botão de remover ultimo item
img3 = PhotoImage(file=f"janela/REMOVER.png")
remover = Button(
    image=img3,
    borderwidth=0,
    highlightthickness=0,
    command=remover_item,
    relief="flat")

remover.place(
    x=240, y=542,
    width=178,
    height=32)

# Crição da caixa de texto
caixa_texto = Text(
    bd=0,
    bg="#ffffff",
    highlightthickness=0)

caixa_texto.place(
    x=240, y=420,
    width=578,
    height=114)

# Campo da entrada de informação sobre o codigo do produto
codigo_produto = Entry(
    bd=0,
    bg="#ffffff",
    highlightthickness=0)

codigo_produto.place(
    x=438, y=224,
    width=280,
    height=25)

# Campo da entrada de informação sobre quantidade
quantidade_prod = Entry(
    bd=0,
    bg="#ffffff",
    highlightthickness=0)

quantidade_prod.place(
    x=438, y=272,
    width=280,
    height=25)

# Menu de formas de pagamento
opcoes = ["Pix", "Dinheiro", "Cartão"]

pagamento = ttk.Combobox(state='readonly', values=opcoes)
pagamento.place(x=438, y=320,
                width=280,
                height=25)

window.resizable(False, False)
window.mainloop()

cursor.close()
conexao.close()
