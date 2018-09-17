# -*- coding: utf-8 -*-
"""
Created on Mon Sep 17 14:21:26 2018

@author: marcos.souto
"""
"""
Esta aplicação pegas os dados da tabela de afastados e atestados e os organizam de forma que consiga em uma tabela fato de maeira que se consiga saber o numero de dias perdidos por mês
 Lembrando que a planilha de ser a primeira entre as demais e nela a de de haver obritariamentes algumas colunas com as seguintes identificações:
     dtInicio = esta coluna e obrigatoria e deve conter a data de inicio da ausencia
     dtFim = Coluna obrigatorio e deve conter a data fim da ausencia
     CID= coluna obrigatoria no arquivo, mas como o CID não é obrigatorio pode conter alguns registro em branco
     CPF= coluna obrigatoria, este registro sera utilizado como chave
     Tempo= coluna obrigatoria, é utilizada para facilitar o calculo de tempo de ausencia, pode ser em dias, meses, anos, o importante é que a dimenção seja igual em todos os registros
"""
import csv
import xlrd
import tkinter as tk
from tkinter import *
from functools import partial
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
import datetime
from datetime import timedelta
from datetime import datetime

#Função abre caixa de dialogo para localizar aquivo .xls*
def fun_Caminho():
        #---Abre a caixa de dialogo e atribui o caminho a string ---------
    caminho = filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("Excel files","*.xls*"),("all files","*.*")))
    return caminho

#Carrega o DataFrame com as informações do arquivo e renomeia as colunas        
def fun_Carrega_DF_Atestados(dfAtest):
    df=pd.DataFrame(dfAtest)
#------Converte as colunas em datatime ------------------------
    df['dtInicio']=pd.to_datetime(df['dtInicio'].squeeze().tolist(),format="%d/%m/%Y")
    df['dtFim']=pd.to_datetime(df['dtFim'].squeeze().tolist(),format="%d/%m/%Y")

    dtInicio=df.dtInicio.min()
    dtFim=df.dtFim.max()
    return df

#Cria uma tabela calendario auxiliar, utilizando a menor data de inicio da ausencia e a maior data fim da ausencia
def fun_Tabela_Calendario(Tempo,strdtInicio):
    dCalendario=[]
    for dtBase in range(1, Tempo):
        dCalendario.append(strdtInicio + timedelta(days=dtBase))
    
    dfCalendario=pd.DataFrame(dCalendario)
    dfCalendario.columns=["dtBase"]
    return dfCalendario


#Cria uma tabela calendario auxiliar, utilizando a menor data de inicio da ausencia e a maior data fim da ausencia
class Application(Frame):
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.place()
        self.grid()
        
        self.lb2=Label()
        self.lb2.configure(text="Qual o nome que deve ser salvo o arquivo?", font="Verdana,25", bg="white")
        self.lb2.place(x=5,y=35)
        
        self.txt1=Entry()
        self.txt1.configure(width=50, text="Exemplo- AtestadosGSM")
        self.txt1.place(x=8,y=70)
        
        self.btArquivo=Button()
        self.btArquivo.configure(width=10, text="arquivo", font="Wingdings, 10",command=lambda:self.fun_inicio())
        self.btArquivo.place(x=320,y=60)
        
        self.lbStatus=Label()
        self.lbStatus.configure(text="")
        self.lbStatus.place(x=10,y=100)
        
    
    def fun_inicio(self):
        strCaminho = fun_Caminho() # Chama a função que localiza o caminho do arquivo
        self.lbStatus.configure()

        if strCaminho !='': #---Verifica se algum arquivo foi selecionado---
            dAtestados = pd.read_excel(strCaminho, sheet_name = "Geral")


        #------------------------------------------------------------------------------------------------------
        #   Cria a tabela fato Atestado com as datas dia a dia
        #   Cria uma tabela fazendo um loop entre a tabela calendario e a tabela de ausencia
        tbAtestado = fun_Carrega_DF_Atestados(dAtestados)
        dtInicio = tbAtestado.dtInicio.min()
        dtFim = tbAtestado.dtFim.max()
        nDias = abs((dtFim - dtInicio).days)
        
        tbCalendario = fun_Tabela_Calendario(nDias, dtInicio)
        lista =[]
        for row1 in tbCalendario.itertuples():
            for row2 in tbAtestado.itertuples():
                if row1.dtBase >= row2.dtInicio and row1.dtBase <= row2.dtFim:
                    lista.append([row1.dtBase, row2.dtInicio, row2.dtFim, row2.CID, row2.RE, row2.Tempo])
        #--------------------------------------------------------------------------------------------------------
        #   Carrega o DataFrame com a lista da tabela fato de ausencias
        dfAtest = pd.DataFrame(lista)
        dfAtest.columns=['dtBase','dtInicio','dtFim','CID','RE','Tempo']                          
        #   Salva a tabela em um arquivo CSV 
        strArquivo = str(self.txt1.get())
        if strArquivo !='':
            NomeArquivo= str(r"C:\Users\marcos.souto\Desktop\Diversos\'") + str(self.txt1.get())
            dfAtest.to_csv( NomeArquivo +".csv" ,sep=',')
            self.lbStatus.configure(text="Arquivo gerado com sucesso")
        else:
                self.lbStatus.configure(text="Necessario escolher o nome do para o arquivo gerado.")
            
            #--------------------------------------------------------------------------------------------------------
            
root = tk.Tk()
#Display(root)
root.title("Criação de Tabela Fato")
root.resizable(height=FALSE,width=FALSE) # não deixa maximizar e nem alterar o frame
root.geometry("450x150+200+150")
root.configure(background="white") 
app=Application(master=root)
app.mainloop()
