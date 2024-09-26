import time
import pyautogui as auto
import pandas as pd
import pyperclip
import os
import tkinter as tk
from tkinter import filedialog, mainloop
from tkinter import messagebox
from pyexpat.errors import messages


class Tela:

    def iniciarTela(self):
        janela = tk.Tk()
        janela.geometry("550x200")
        janela.title("Cupons Fiscais")
        janela.resizable(False,False)
        self.entry_buscar = tk.Entry(janela,width=50)
        self.entry_buscar.grid(row=0, column=0, padx=10, pady=10)
        self.button_buscar = tk.Button(janela,text="Buscar",width=10,command=self.buscar_arquivo)
        self.button_buscar.grid(row=0, column=1, padx=10, pady=10)
        botao_start = tk.Button(janela,text="START", width=50, command=self.iniciarProcesso)
        botao_start.grid(row=1, column=0, padx=10, pady=10)
        janela.mainloop()

    def buscar_arquivo(self):
        file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx *.xls")])  # Filtros para arquivos Excel
        self.entry_buscar.delete(0, tk.END)  # Limpa o campo de entrada
        self.entry_buscar.insert(0, file_path)  # Insere o caminho no campo de entrada

    def iniciarProcesso(self):
        try:
            df = pd.read_excel(self.entry_buscar.get())
            mensagem = messagebox.showinfo(title="Aviso",
                                           message="Ao clicar em OK você terá 20 segundos para entrar na tela da Personalcard.")
            time.sleep(20)
        except Exception:
            messagebox.showinfo(title="Aviso",message="Não foi possível ler o arquivo passado")
            return
        coluna = df[['CHAVESAT', 'CGC']]
        auto.scroll(300)
        for dados in coluna.values:
            auto.moveTo(1422, 696, duration=2)
            auto.click()
            auto.hotkey('ctrl', 'a')
            auto.hotkey('backspace')
            cpf = dados[1]
            pyperclip.copy(cpf)
            auto.hotkey('ctrl', 'v')
            auto.moveTo(746, 795, duration=2)
            auto.click()
            auto.moveTo(1730, 965, duration=2)
            auto.click()
            auto.moveTo(763, 737, duration=2)
            auto.click()
            chave = str(dados[0])
            pyperclip.copy(chave)
            auto.hotkey('ctrl', 'v')
            auto.moveTo(1540, 744, duration=2)
            auto.click()
            auto.moveTo(959, 739, duration=2)
            auto.click()
        messagebox.showinfo(title="CONCLUIDO", message="Script Concluido com sucesso!!!")


tela = Tela()
tela.iniciarTela()


















