import xmltodict
import os
from datetime import datetime
from re import sub
import pandas as pd
from tkinter.filedialog import askdirectory, asksaveasfilename
from tkinter import *
from tkinter import ttk
import time
from threading import Thread
""" import json """

def tela():
    tela = Tk()
    tela.title("XML to XLSX")
    tela.geometry("300x200")

    progressbar = ttk.Progressbar(tela)
    progressbar.place(x=30, y=60, width=200)

    explore_button = Button(text="Find", command=lambda: Thread(target=lambda: path(progressbar, tela)).start())
    explore_button.place(x=30, y=20)
    
    tela.mainloop()
    
def pegar_infos(arquivo, valores, path):
    """ print(f"Pegou as informações {arquivo}") """
    with open(f"{path}/{arquivo}", "rb") as arquivo_xml:
        dic_arquivo = xmltodict.parse(arquivo_xml)
        """ try: """
        if 'nfeProc' in dic_arquivo:
            infos_nf = dic_arquivo["nfeProc"]["NFe"]["infNFe"]
        else:
            infos_nf = dic_arquivo["NFe"]["infNFe"]
        numero_nota = infos_nf['ide']['nNF']
        if 'dhEmi' in infos_nf['ide']:
            data_emissao_sf = infos_nf['ide']['dhEmi']
        else:
            data_emissao_sf = infos_nf['ide']['dEmi']
        chave_acesso = infos_nf['@Id']
        cnpj_emissor = infos_nf['emit']['CNPJ']
        empresa_emissora = infos_nf['emit']['xFant']
        nome_cliente = infos_nf['dest']['xNome']
        valor_total = infos_nf['total']['ICMSTot']['vProd']
        valor_frete = infos_nf['total']['ICMSTot']['vFrete']

        #Formatação da Hora:
        iso_date = data_emissao_sf
        dt = datetime.fromisoformat(iso_date)
        data_emissao = dt.strftime('%d/%m/%Y')

        #Formatação da Chave de Acesso (44 digitos)
        chave_acesso = sub('[A-Za-z]', '', chave_acesso)

        #Formatação do Preço:
        valor_total = f"R${valor_total}"
        valor_frete = f"R${valor_frete}"

        valores.append([data_emissao, empresa_emissora, numero_nota, chave_acesso, cnpj_emissor, nome_cliente, valor_total, valor_frete])
        """ except Exception as e:
            print(e)
            print(json.dumps(dic_arquivo,  indent=4))  """
        
def path(progressbar, tela):
    path = askdirectory(title="Select a Folder")
    lista_arquivos = os.listdir(path)

    colunas = ["data_emissao", "empresa_emissora", "numero_nota", "chave_acesso", "cnpj_emissor", "nome_cliente", "valor_total", "valor_frete"]
    valores = []
    carregar(progressbar)

    for arquivo in lista_arquivos:
        pegar_infos(arquivo, valores, path)
    
    arquivo = asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], initialfile="NotasFiscais.xlsx", title="Selecione onde salvar o arquivo Excel")
    tabela = pd.DataFrame(columns=colunas, data=valores)
    tabela.to_excel(arquivo, index=False)
    time.sleep(1)
    tela.destroy()

def carregar(progressbar):
    for i in range(101):
        time.sleep(0.01)  
        progressbar["value"] = i
        progressbar.update()


if __name__ == "__main__":
    tela()
