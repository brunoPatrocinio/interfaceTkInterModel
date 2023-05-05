import tkinter as tk
import pandas as pd
from tkinter import ttk
from tkinter.filedialog import askopenfilename
from tkcalendar import DateEntry
from datetime import datetime
import requests #lib para fazer requisições
import numpy as np

requisicoes = requests.get("https://economia.awesomeapi.com.br/json/all")
dicionarioMoedas = requisicoes.json() #trazendo da api para dentro do dicionario
caminhoPad = ""

listaMoedas = list(dicionarioMoedas.keys()) #adiciona os valores do dicionario na lista, que vai na combobox

#funções
def pegaCotacao():
    moeda = comboBox_selecionar_moeda.get()
    dataCotacao = calendarioMoeda.get()
    ano = dataCotacao[-4:]
    mes = dataCotacao[3:5]
    dia = dataCotacao[:2]
    try:
        link = f"https://economia.awesomeapi.com.br/json/daily/{moeda}-BRL/?start_date={ano}{mes}{dia}&end_date={ano}{mes}{dia}"
        requisicaoMoeda = requests.get(link)
        cotacao = requisicaoMoeda.json()
        valorMoeda = cotacao[0]["bid"]
        labelTextoCotacao["text"] = f"A cotação da moeda {moeda} no dia {dataCotacao} foi de: R${valorMoeda}"
    except:
        labelTextoCotacao["text"] = "Erro, ou não existe cotação para esse dia ainda."

def selecionarArquivo():
    caminhoArquivo = askopenfilename(title="Selecione o arquivo de moeda...")
    varCaminhoArquivo.set(caminhoArquivo)
    if caminhoArquivo:
        labelArquivoSelecionado["text"] = f"Arquivo selecionado {caminhoArquivo}"

def atualizarCotacoes():
    try:
        dt = pd.read_excel(varCaminhoArquivo.get())
        moedas = dt.iloc[:, 0]
        dataInicial = calendarioDataInicial.get()
        dataFinal = calendarioDataFinal.get()
        anoInicial = dataInicial[-4:]
        mesInicial = dataInicial[3:5]
        diaInicial = dataInicial[:2]

        anoFinal = dataInicial[-4:]
        mesFinal = dataInicial[3:5]
        diaFinal = dataInicial[:2]
        for moeda in moedas:
            link = f"https://economia.awesomeapi.com.br/json/daily/{moeda}-BRL/?" \
                   f"start_date={anoInicial}{mesInicial}{diaInicial}&end_date={anoFinal}{mesFinal}{diaFinal}"
            requisicaoMoeda = requests.get(link)
            cotacoes = requisicaoMoeda.json()
            for cotacao in cotacoes:
                timestamp = int(cotacao["timestamp"])
                bid = float(cotacao["bid"])
                data = datetime.fromtimestamp(timestamp)

                data = data.strftime('%d/%m/%Y') #transf. num texto com cara de data
                if data not in dt:
                    dt[data] = np.nan

                dt.loc[dt.iloc[:, 0] == moeda, data] = bid
        dt.to_excel("moedas.xlsx")


        labelAtualizarCotacoes['text'] = "Arquivo Atualizado com Sucesso"

        print(dt)
    except:
        labelAtualizarCotacoes['text'] = "Selecione um arquivo Excel no Formato Correto"

janela = tk.Tk()
janela.title("Ferramenta de Cotação de Moedas")

label_cotacaomoeda = tk.Label(text="Cotação de 1 moeda específica", borderwidth=2, relief="solid")
label_cotacaomoeda.grid(row=0, column=0, padx=10, pady=10, sticky="nswe", columnspan=3)

label_selecionarMoeda = tk.Label(text="Selecionar Moeda:", anchor="e")
label_selecionarMoeda.grid(row=1, column=0, padx=10, pady=10, sticky="nsew", columnspan=2)

#Combobox
comboBox_selecionar_moeda = ttk.Combobox(values=listaMoedas)
comboBox_selecionar_moeda.grid(row=1, column=2, padx=10, pady=10, sticky="nsew")

label_selecionarDiaCot = tk.Label(text="Selecione o dia que deseja obter a cotação", anchor="e")
label_selecionarDiaCot.grid(row=2, column=0, padx=10, pady=10, sticky="nsew", columnspan=2)

calendarioMoeda = DateEntry(year=2023, locale="pt_br") #calendario com formato em portugues brasil
calendarioMoeda.grid(row=2, column=2, pady=10, padx=10, sticky="nswe", columnspan=2)

labelTextoCotacao = tk.Label(text="")
labelTextoCotacao.grid(row=3, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")

#botão pegar cotação
botaPegaCotacao = tk.Button(text="Obter Cotação", command=pegaCotacao)
botaPegaCotacao.grid(row=3, column=2, padx=10, pady=10, sticky="nsew")

#selecionar varias moedas
label_variasMoedas = tk.Label(text="Cotação de múltiplas moedas", borderwidth=2, relief="solid")
label_variasMoedas.grid(row=4, column=0, padx=10, pady=10, sticky="nswe", columnspan=3)

labelSelecionarArquivo = tk.Label(text="Selecione um arquivo em Excel com as Moedas na Coluna A")
labelSelecionarArquivo.grid(row=5, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")

varCaminhoArquivo = tk.StringVar()

botaoSelecionarArquivo = tk.Button(text="Clique para Selecionar", command=selecionarArquivo)
botaoSelecionarArquivo.grid(row=5, column=2, padx=10, pady=10, sticky="nsew")

labelArquivoSelecionado = tk.Label(text="Nenhum Arquivo Selecionado.", anchor="e")
labelArquivoSelecionado.grid(row=6, column=0, columnspan=3, padx=10, pady=10, sticky="nsew")

#data inicial
labelDataInicial = tk.Label(text="Data Inicial:", anchor="e")
labelDataInicial.grid(row=7, column=0, padx=10, pady=10, sticky="nsew")

#data final
labelDataFinal = tk.Label(text="Data Final:", anchor="e")
labelDataFinal.grid(row=8, column=0, padx=10, pady=10, sticky="nsew")

calendarioDataInicial = DateEntry(year=2023, locale="pt_br")
calendarioDataFinal = DateEntry(year=2023, locale="pt_br")
calendarioDataInicial.grid(row=7, column=1, padx=10, pady=10, sticky="nsew")
calendarioDataFinal.grid(row=8, column=1, padx=10, pady=10, sticky="nsew")

#botão atualizar cotações
botaoAtualizarCotacoes = tk.Button(text="Atualizar Cotações", command=atualizarCotacoes)
botaoAtualizarCotacoes.grid(row=9, column=0, padx=10, pady=10, sticky="nsew")

labelAtualizarCotacoes = tk.Label(text="")
labelAtualizarCotacoes.grid(row=9, column=1, columnspan=2, padx=10, pady=10, sticky="nsew")

#fechar
botaoFechar = tk.Button(text="Fechar", command=janela.quit)
botaoFechar.grid(row=10, column=2, padx=10, pady=10, sticky="nsew")

janela.mainloop()
