#Exportação de Anilhas, by GabsTIH

<<<<<<< HEAD
version = "v1.1"

#Importação das bibliotecas utilizadas
import pandas as pd
=======
#Importação das bibliotecas utilizadas
import pandas as pd
import os
>>>>>>> f29a63cea8e801fc59041b56e47f48d7b2f4f1eb
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

<<<<<<< HEAD
def getCamin():
    """
    Pede ao usuário o caminho do arquivo que ele deseja que seja feita a exportação
    (Exemplo: C:Users/Fulano/Desktop/anilhas.xlsx)
    """

    global caminho
    caminho = filedialog.askopenfilename(
        filetypes=[('Planilhas do Excel', '*xls *xlsx *XLS *XLSX')]
=======
#Pede o caminho do arquivo pro usuário (Exemplo: C:Users/Fulano/Desktop/anilhas.xlsx)

def getCamin():
    global caminho
    caminho = filedialog.askopenfilename(
        filetypes=[('Planilhas do Excel', '*xls *xlsx')]
>>>>>>> f29a63cea8e801fc59041b56e47f48d7b2f4f1eb
    )
    caminn.configure(text = f'Caminho selecionado: {caminho}')

def exportacao():
<<<<<<< HEAD
    """
    Realiza o processo de exportação utilizando o arquivo excel selecionado pelo usuário
    e retorna com um arquivo novo com as duas colunas da tabela de anilhas.
    """
    #Aviso caso o usuário não selecione um arquivo.
    if 'caminho' not in globals() or caminho == "":
        warning1.configure(text = "Por favor, selecione um arquivo primeiro!")
        return
    else:
    #Pergunta ao usuário se ele deseja continuar o processo de exportação após selecionar um arquivo.
=======
    if 'caminho' not in globals() or caminho == "":
        aviso1.configure(text = "Por favor, selecione um arquivo primeiro!")
        return
    else:
>>>>>>> f29a63cea8e801fc59041b56e47f48d7b2f4f1eb
        confirme = messagebox.askyesno("Confirmar exportação", "Você tem certeza de que deseja exportar este arquivo?")
        if not confirme:
            return
        else:
            try:
                # Verifica se o arquivo é de fato uma planilha
<<<<<<< HEAD
                if '.xlsx' in caminho.lower():
                    table = pd.read_excel(caminho, engine='openpyxl')
                elif '.xls' in caminho.lower():
                    table = pd.read_excel(caminho, engine='xlrd')
                else:
                    warning1.configure(text="Arquivo não suportado!")
                    return
            except Exception as e:
                warning1.configure(text="Erro ao ler o arquivo.")
        warning1.configure(text='')
        nomeExportado = nameExp.get()
    if nomeExportado == "" or nomeExportado == "()":
        warning1.configure(text = "Por favor digite um nome pro arquivo exportado!")
        return
    else:
        warning1.configure(text='')
=======
                if '.xlsx' in caminho:
                    table = pd.read_excel(caminho, engine='openpyxl')
                elif 'xls' in caminho:
                    table = pd.read_excel(caminho, engine='xlrd')
                else:
                    aviso1.configure(text="Arquivo não suportado!")
                    return
            except Exception as e:
                aviso1.configure(text="Erro ao ler o arquivo.")
        aviso1.configure(text='')
        nomeExportado = nomeExp.get()
    if nomeExportado == "" or nomeExportado == "()":
        aviso1.configure(text = "Por favor digite um nome pro arquivo exportado!")
        return
    else:
        aviso1.configure(text='')
>>>>>>> f29a63cea8e801fc59041b56e47f48d7b2f4f1eb
        global exportado
        exportado = table.iloc[12:, [5, 13]]
        print(exportado)
        exportado.to_excel(f'{nomeExportado}.xlsx', index = False, header = False)
        status.configure(text=f"Arquivo {nomeExportado} exportado !")

<<<<<<< HEAD
bg_color = "#222831"
text_color = "#EEEEEE"
fonte = ['Arial', 14]

window = tk.Tk()
window.title("Anilha Exporter | Desenvolvido por Gabriel Barbosa")
window.geometry("500x800")
window.resizable(False, False)
window.configure(bg=bg_color)

title = tk.Label(
                    window,
                    text= f"Anilha Exporter {version}",
                    font=('Arial', 34),
                    bg=bg_color,
                    fg=text_color,
=======
cor_fundo = "#dce0e2"

window = tk.Tk()
window.title("Exportador de Anilhas | By Gabs")
window.geometry("500x800")
window.resizable(False, False)
window.configure(bg=cor_fundo)

boasvindas = tk.Label(
                    window,
                    text="Boas-vindas !",
                    font=('Arial', 16),
                    bg=cor_fundo,
>>>>>>> f29a63cea8e801fc59041b56e47f48d7b2f4f1eb
                    anchor="center",
                    justify="center"
)

<<<<<<< HEAD
title.grid(row=0, columnspan=2, padx= 15, pady=20, sticky="nsew")

order1 = tk.Label(window, text="Por favor escolha o caminho abaixo", font=(fonte), bg=bg_color, fg=text_color)
warning1 = tk.Label(window, text='', bg=bg_color, fg=text_color, font=('Arial', 10, "underline"))
order1.grid(row=1, column=0, columnspan=2, padx=15, pady=20)
warning1.grid(row=8, column=0, columnspan=2, padx=15, pady=20)

camin = tk.Button(window, text="Selecionar Arquivo", command=getCamin, bg="#76ABAE")
caminn = tk.Label(window, text='', font=(fonte), wraplength=300, bg=bg_color, fg=text_color)
camin.grid(row=2, column=0, columnspan=2, padx=15, pady=20)
caminn.grid(row=3, column=0, padx=15, pady=20)

nameExp = tk.Entry(window, bg="white")
ordem2 = tk.Label(window, text='Digite o nome do arquivo exportado:',font=(fonte), bg=bg_color, fg=text_color)
ordem2.grid(row=4, column=0, columnspan=2, padx=15, pady=20)
nameExp.grid(row=5, column=0, columnspan=2, padx=15, pady=20)

confirm = tk.Button(window, text="Exportar", command=exportacao, bg="#76ABAE")
confirm.grid(row=6, column=0, columnspan=2, padx=15, pady=20)

status = tk.Label(window, text='', bg=bg_color, fg=text_color)
status.grid(row=7, column=0, columnspan=2, padx=15, pady=20)

footer = tk.Label(window, text=f'{version} - Desenvolvido por Gabriel Barbosa', font=("Arial", 9), bg=bg_color, fg="gray")
footer.grid(row=99, column=0, columnspan=2, pady=20)
=======
boasvindas.grid(row=0, columnspan=2, padx= 15, pady=20, sticky="nsew")

ordem1 = tk.Label(window, text="Por favor escolha o caminho abaixo", font=('Arial', 12), bg=cor_fundo)
aviso1 = tk.Label(window, text='', bg=cor_fundo)
ordem1.grid(row=1, column=0, columnspan=2, padx=15, pady=20)
aviso1.grid(row=8, column=0, columnspan=2, padx=15, pady=20)

camin = tk.Button(window, text="Selecionar Arquivo", command=getCamin, bg="#d3dee4")
caminn = tk.Label(window, text='', font=('Arial', 12), wraplength=300, bg=cor_fundo)
camin.grid(row=2, column=0, columnspan=2, padx=15, pady=20)
caminn.grid(row=3, column=0, padx=15, pady=20)

nomeExp = tk.Entry(window, bg="white")
ordem2 = tk.Label(window, text='Digite o nome do arquivo exportado:',font=('Arial', 12), bg=cor_fundo)
ordem2.grid(row=4, column=0, columnspan=2, padx=15, pady=20)
nomeExp.grid(row=5, column=0, columnspan=2, padx=15, pady=20)

confirmar = tk.Button(window, text="Exportar", command=exportacao, bg="#d5f0e3")
confirmar.grid(row=6, column=0, columnspan=2, padx=15, pady=20)

status = tk.Label(window, text='', bg=cor_fundo)
status.grid(row=7, column=0, columnspan=2, padx=15, pady=20)

preview = tk.Label(window, text='', bg=cor_fundo, wraplength=200)
preview.grid(row=4, column=2, padx=15, pady=20)

rodape = tk.Label(window, text="v1.0 - Desenvolvido por Gabs para Grupo BTM", font=("Arial", 8), bg=cor_fundo, fg="gray")
rodape.grid(row=99, column=0, columnspan=2, pady=20)
>>>>>>> f29a63cea8e801fc59041b56e47f48d7b2f4f1eb

window.columnconfigure(0, weight=1)
window.mainloop()