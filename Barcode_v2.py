from barcode import Code128
from barcode.writer import ImageWriter
import os
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
from PIL import Image, ImageTk
import openpyxl

def entrada_e_gerar_codigo():
    global serial
    serial = caixatexto.get()

    # Definição do código de barras
    number = serial

    # Gerando e renderizando imagem
    my_barcode = Code128(number, writer=ImageWriter())
    savePath = os.path.join(os.environ['USERPROFILE'], "Desktop", serial)
    my_barcode.save(savePath)

    messagebox.showinfo("Código de Barras", "Código de barras gerado com sucesso!")

def gerar_codigos_excel():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])

    if file_path:
        try:
            workbook = openpyxl.load_workbook(file_path)
            sheet = workbook.active

            output_folder = os.path.join(os.environ['USERPROFILE'], "Desktop", "CodigosDeBarras")
            os.makedirs(output_folder, exist_ok=True)

            for row in sheet.iter_rows(min_row=2, values_only=True):
                serial = row[0]
                my_barcode = Code128(serial, writer=ImageWriter())
                savePath = os.path.join(output_folder, serial)
                my_barcode.save(savePath)

            messagebox.showinfo("Códigos de Barras", "Códigos de barras gerados com sucesso!")

        except Exception as e:
            messagebox.showerror("Erro", str(e))

# Interface gráfica
janela = tk.Tk()
janela.title("Gerador de Código de Barras")
janela.geometry("400x250")
janela.configure(background="#34495E")  # Fundo em azul Blockbit
janela.iconbitmap("C:/Users/guilh/Desktop/Cursos/Algoritmos/Exercícios/BarcodeGen/Barras.ico")  # Adicione o caminho para o arquivo .ico

# Estilo para os elementos da interface
style = ttk.Style()
style.configure("TLabel", background="#34495E", foreground="white", font=("Helvetica", 12))
style.configure("TButton", background="#2ECC71", foreground="black", font=("Helvetica", 12, "bold"))

texto_orientacao = ttk.Label(janela, text="Insira abaixo o serial do equipamento:", anchor="center")
texto_orientacao.place(x=40, y=10, width=320, height=40)

caixatexto = ttk.Entry(janela, font=("Helvetica", 12))
caixatexto.place(x=40, y=60, width=320, height=30)

# Botão estilizado com ícone
image = Image.open("C:/Users/guilh/Desktop/Cursos/Algoritmos/Exercícios/BarcodeGen/save.ico")  # Adicione o caminho para o ícone
image = image.resize((20, 20), Image.ANTIALIAS)
icon = ImageTk.PhotoImage(image)
botao = ttk.Button(janela, text="Salvar e Gerar Código", image=icon, compound="left", command=entrada_e_gerar_codigo)
botao.place(x=40, y=110, width=320, height=40)

# Botão para gerar códigos a partir de um arquivo Excel
botao_gerar_excel = ttk.Button(janela, text="Gerar Códigos a partir do Excel", command=gerar_codigos_excel)
botao_gerar_excel.place(x=40, y=170, width=320, height=40)

janela.mainloop()





