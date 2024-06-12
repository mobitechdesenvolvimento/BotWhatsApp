import tkinter as tk
from tkinter import filedialog, messagebox
import openpyxl
from urllib.parse import quote
import webbrowser
import time
import pyautogui


def enviar_mensagens():
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx")])
    if not file_path:
        messagebox.showerror("Erro", "Nenhuma planilha selecionada.")
        return

    try:
        # Abrir WhatsApp Web
        webbrowser.open('https://web.whatsapp.com/')
        time.sleep(30)  # Aguarda 30 segundos para que o WhatsApp Web carregue

        # Ler planilha e guardar informações sobre nome e telefone
        workbook = openpyxl.load_workbook(file_path)
        pagina_demandas = workbook['cadastros']

        for linha in pagina_demandas.iter_rows(min_row=2):
            Nome = linha[6].value
            Telefone = linha[20].value

            if not Nome or not Telefone:
                continue

            # Formatar telefone: remover caracteres especiais e adicionar o código do país
            Telefone = Telefone.replace('(', '').replace(
                ')', '').replace('-', '').replace(' ', '')
            if not Telefone.startswith('55'):
                Telefone = '55' + Telefone

            # Mensagem
            mensagem = f'Olá {
                Nome}, conto com seu voto! https://www.instagram.com/dadoseleitorais'
            # Criar link personalizado do WhatsApp
            link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={
                Telefone}&text={quote(mensagem)}'
            webbrowser.open(link_mensagem_whatsapp)
            time.sleep(10)  # Ajuste o tempo conforme necessário

            # Pressionar Enter para enviar a mensagem
            pyautogui.press('enter')
            time.sleep(6)  # Pausa para garantir que a mensagem seja enviada

            # Fechar a aba atual
            pyautogui.hotkey('ctrl', 'w')
            time.sleep(2)

        messagebox.showinfo("Sucesso", "Mensagens enviadas com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")


# Configurar a interface gráfica
root = tk.Tk()
root.title("Envio Automático de Mensagens no WhatsApp")

frame = tk.Frame(root)
frame.pack(pady=20)

botao_abrir = tk.Button(frame, text="Carregar Planilha",
                        command=enviar_mensagens)
botao_abrir.pack(pady=10)

root.mainloop()
