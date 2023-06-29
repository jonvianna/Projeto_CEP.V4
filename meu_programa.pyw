import pandas as pd
import tkinter as tk
from tkinter import messagebox
import os
import sys
# import win32gui

# Impedido a janela do CMD abrir junto quando executado o app
# win32gui.ShowWindow(win32gui.GetForegroundWindow(), 0)


# Lendo o arquivo Excel
file_path = os.path.join(sys._MEIPASS, "PlanilhaCidadesCep.xlsx") if hasattr(sys, '_MEIPASS') else "PlanilhaCidadesCep.xlsx"
df = pd.read_excel(file_path)

# Função para verificar o CEP
def verificar_cep():
    # Obtendo o CEP digitado pelo usuário
    cep = entrada_cep.get()

    # Verificando se o CEP possui 8 dígitos
    if len(cep) != 8:
        messagebox.showinfo("Erro", "Digite um CEP válido")
        entrada_cep.select_range(0, tk.END)
        return

    # Verificando se o CEP está dentro do range de algum registro na tabela
    if cep[0] in ['5', '6', '7']:
        resultado.set("Fora da área de coleta! NÃO TRABALHAMOS COM ESTA REGIÃO!!")
        messagebox.showinfo("Resultado", "Fora da área de coleta! \nNÃO TRABALHAMOS COM ESTA REGIÃO!!")
        entrada_cep.select_range(0, tk.END)
    elif cep[0] == '4':
        for i, row in df.iterrows():
            if row['Cep inicio'] <= int(cep) <= row['Cep fim']:
                cidade = row['Cidades']
                resultado.set(cidade + " possui coleta")
                messagebox.showinfo("Resultado", cidade + " - Possui coleta")
                entrada_cep.select_range(0, tk.END)
                break
        else:
            resultado.set("Fora da área de coleta! NÃO TRABALHAMOS COM ESTA REGIÃO!!")
            messagebox.showinfo("Resultado", "Fora da área de coleta! \nNÃO TRABALHAMOS COM ESTA REGIÃO!!")
            entrada_cep.select_range(0, tk.END)
    else:
        for i, row in df.iterrows():
            if row['Cep inicio'] <= int(cep) <= row['Cep fim']:
                cidade = row['Cidades']
                resultado.set(cidade + " possui coleta")
                messagebox.showinfo("Resultado", cidade + " - Possui coleta")
                entrada_cep.select_range(0, tk.END)
                break
        else:
            resultado.set("Fora da área de coleta!")
            messagebox.showinfo("Resultado", "Fora da área de coleta!")
            entrada_cep.select_range(0, tk.END)

# Função para encerrar o programa
def fechar_janela():
    janela.quit()

# Criando a janela principal
janela = tk.Tk()
janela.title("Verificador de CEP")
janela.geometry("300x70")

# Criando o rótulo e a entrada para o CEP
tk.Label(janela, text="Digite um CEP:").pack()
entrada_cep = tk.Entry(janela)
entrada_cep.pack()

# Criando o botão para verificar o CEP
botao_verificar = tk.Button(janela, text="Verificar", command=verificar_cep)
botao_verificar.pack()

# Vinculando a tecla Enter à função de verificação do CEP
entrada_cep.bind('<Return>', lambda event: verificar_cep())

# Criando o rótulo para exibir o resultado
resultado = tk.StringVar()

# Definindo a função para encerrar o programa quando a janela for fechada
janela.protocol("WM_DELETE_WINDOW", fechar_janela)

# Iniciando o loop principal do Tkinter
janela.mainloop()
