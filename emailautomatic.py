# Importa a biblioteca necessária para lidar com o Microsoft Outlook
import win32com.client as win32
# Importa a biblioteca tkinter para a interface gráfica
import tkinter as tk
# Importa o módulo filedialog da biblioteca tkinter para lidar com caixas de diálogo de arquivos

from tkinter import filedialog

# Função para enviar o e-mail
def enviar_email():
    # Cria uma instância do Outlook
    outlook = win32.Dispatch('outlook.application')
    # Cria um novo e-mail
    mail = outlook.CreateItem(0)
    # Preenche o campo "Para" com o valor inserido pelo usuário
    mail.to = entry_para.get()
    # Preenche o campo "Assunto" com o valor inserido pelo usuário
    mail.Subject = entry_assunto.get()
    # Preenche o corpo do e-mail com o valor inserido pelo usuário
    mail.HTMLBody = entry_corpo.get("1.0", "end-1c")

    # Anexa arquivos, se especificado pelo usuário
    anexo = entry_anexo.get()
    if anexo:
        mail.Attachments.Add(anexo)

    # Envia o e-mail
    mail.Send()

# Função para selecionar um anexo
def selecionar_anexo():
    # Abre uma janela de seleção de arquivo e obtém o nome do arquivo selecionado
    filename = filedialog.askopenfilename()
    # Limpa o campo de anexo e insere o nome do arquivo selecionado nele
    entry_anexo.delete(0, tk.END)
    entry_anexo.insert(0, filename)

# Cria a janela principal da interface gráfica
root = tk.Tk()
root.title("Envio de E-mail")

# Define a cor de fundo como meio transparente e azul claro
root.attributes('-alpha', 0.8)
root.configure(bg='light blue')

# Campos e rótulos para preencher os dados do e-mail (Para, Assunto, Corpo, Anexo)
label_para = tk.Label(root, text="Para:")
label_para.pack()
entry_para = tk.Entry(root)
entry_para.pack()

label_assunto = tk.Label(root, text="Assunto:")
label_assunto.pack()
entry_assunto = tk.Entry(root)
entry_assunto.pack()

label_corpo = tk.Label(root, text="Corpo:")
label_corpo.pack()
entry_corpo = tk.Text(root, height=10, width=50)
entry_corpo.pack()

label_anexo = tk.Label(root, text="Anexo:")
label_anexo.pack()
entry_anexo = tk.Entry(root)
entry_anexo.pack()

# Botão para selecionar um anexo (chama a função selecionar_anexo)
button_selecionar_anexo = tk.Button(root, text="Selecionar Anexo", command=selecionar_anexo)
button_selecionar_anexo.pack()

# Botão para enviar o e-mail (chama a função enviar_email)
button_enviar = tk.Button(root, text="Enviar", bg="green", command=enviar_email)
button_enviar.pack()

# Inicia o loop principal da interface gráfica, mantendo-a ativa e aguardando interações do usuário
root.mainloop()