# Importar bibliotecas necessárias
# pip install tkinter pandas webbrowser pyautogui ttkbootstrap

import tkinter as tk
from tkinter import ttk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.style import Style    
from urllib.parse import quote
from time import sleep
from datetime import datetime
import pandas as pd
import webbrowser
import threading
import pyautogui 


# Criar uma classe para a GUI
class CourseOfferGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Disparador de Mensagem 2.0")
        self.root.geometry("500x600")


        #TÍTULO
        app = self.root
        label = ttk.Label(app, text='Responda para começarmos!')
        label.pack(pady=10)
        label.config(font=('Arial', 20, 'bold'))
        style = Style(theme='superhero')

        # Criar campos de entrada
        self.course_label = ttk.Label(root, text="Curso:")
        self.course_label.pack()
        self.course_entry = ttk.Entry(root, width=30)
        self.course_entry.pack()

        self.partner_label = ttk.Label(root, text="Instituição Parceira:")
        self.partner_label.pack()
        self.partner_entry = ttk.Entry(root, width=30)
        self.partner_entry.pack()

        self.period_label = ttk.Label(root, text="Período:")
        self.period_label.pack()
        self.period_entry = ttk.Entry(root, width=30)
        self.period_entry.pack()

        self.schedule_label = ttk.Label(root, text="Horário:")
        self.schedule_label.pack()
        self.schedule_entry = ttk.Entry(root, width=30)
        self.schedule_entry.pack()

        self.minage_label = ttk.Label(root, text="Idade mínima:")
        self.minage_label.pack()
        self.minage_entry = ttk.Entry(root, width=30)
        self.minage_entry.pack()

        self.duration_label = ttk.Label(root, text="Duração:")
        self.duration_label.pack()
        self.duration_entry = ttk.Entry(root, width=30)
        self.duration_entry.pack()

        self.group_label = ttk.Label(root, text="Deseja realizar o envio por grupo?(SIM/NÃO):")
        self.group_label.pack()
        self.group_entry = ttk.Entry(root, width=30)
        self.group_entry.pack()

        self.minrange_label = ttk.Label(root, text="De qual linha devo começar:")
        self.minrange_label.pack()
        self.minrange_entry = ttk.Entry(root, width=30)
        self.minrange_entry.pack()

        self.maxrange_label = ttk.Label(root, text="Até qual linha devo enviar:")
        self.maxrange_label.pack()
        self.maxrange_entry = ttk.Entry(root, width=30)
        self.maxrange_entry.pack()

        # Criar um botao para enviar mensagens
        self.send_button = ttk.Button(root, text="Enviar Mensagens", command=self.start_sending)
        self.send_button.pack(pady=7)

        # Criar um botao para cancelar o código
        self.cancel_button = ttk.Button(root, text="Interromper código", bootstyle=DANGER, command=self.interromper_codigo)
        self.cancel_button.pack(pady=7)

        #CRÉDITOS
        app = self.root
        label = ttk.Label(app, text='developed by: Lucas Ferrari')
        label.pack(pady=5)
        label.config(font=('Arial', 7, 'bold'))

        self.running = False

    def interromper_codigo(self):
        self.running = False

    def start_sending(self):
        if self.running:
            print("O envio de mensagens já está em andamento.")
            return
        self.running = True
        print('Iniciado o envio de mensagens!')
        t = threading.Thread(target=self.send_messages)
        t.start()

    def encontra_categoria(self, curso_de_envio: str) -> list:
        #DEFINE AS CATEGORIAS, ADICIONAR TODOS OS NOVOS CURSOS DA AMTECH
        Dev=["INTRODUÇÃO A LÓGICA DE PROGRAMAÇÃO", "INTRODUÇÃO A ALGORÍTIMOS", "INTRODUÇÃO EM ROBÓTICA", "INTRODUÇÃO A UTILIZAÇÃO DE IAS E CHATBOTS DE FORMA PRODUTIVA"]

        Marketing=["INSTAGRAM PARA NEGÓCIOS", "MARKETING DIGITAL", "CRIAÇÃO DE MÍDIAS PARA REDES SOCIAIS", "REDES SOCIAIS: UM GUIA PRÁTICO  PARA ALAVANCAR SUAS VENDAS"]      

        Design=["INTRODUÇÃO A INOVAÇÃO E DESIGN", "CRIAÇÃO DE MÍDIAS PARA REDES SOCIAIS", "SKETCHUP - SOFTWARE CRIAÇÃO DE MODELOS EM 3D", "IMPRESSORA 3D - BÁSICO"]

        Pacote_Office=["PACOTE OFFICE", "EXCEL(INTERMEDIÁRIO)", "EXCEL(AVANÇADO)"]

        Basicos=["INTRODUÇÃO A INFORMÁTICA", "INFORMÁTICA BÁSICA (TERCEIRA IDADE)"]
        lista_categoria = [Dev, Marketing, Design, Pacote_Office, Basicos]

        Correto = None
        while Correto == None:
            for Categorias in lista_categoria:
                if curso_de_envio in Categorias:
                    Correto = Categorias
        return Correto


    def send_messages(self):
        # Obter entrada do usuário
        curso_de_envio = self.course_entry.get()
        parceiro = self.partner_entry.get()
        periodo = self.period_entry.get()
        horario_do_curso = self.schedule_entry.get()
        data_de_duracao = self.duration_entry.get()
        linhamin = int(self.minrange_entry.get())
        linhamax = int(self.maxrange_entry.get())
        idademin = int(self.minage_entry.get())
        por_grupo = self.group_entry.get()
    

        # Ler o arquivo Excel
        alunos = pd.read_excel('alunos.xlsx')

        # Definir running como True para iniciar o envio de mensagens
        self.running = True

        #pega a string de interesse de cada aluno, transforma cada curso em um item de uma lista.
        #essa range corresponde ao número de linhas, lembrando que vai sempre ler uma linha a menos do que a gente informar.
        for x in range(linhamin, linhamax):

                if not self.running:
                    print('Código interrompido na linha: {0}'.format(x))
                    break
                
                #X refere-se a linha, "Dentre as opções qual curso gostaria de fazer" se trata da coluna a ser lida (ele pega pelo cabeçalho)
                cursos = alunos.loc[x, "Dentre as opções qual curso gostaria de fazer?"]
                lista_cursos = cursos.split(sep=', ')

                if por_grupo == "SIM":
                    categoria = self.encontra_categoria(curso_de_envio)
                    for curso in lista_cursos:
                        if curso in categoria:
                            # Ler planilha e guardar informações nome e telefone
                            nome = alunos.loc[x, 'Nome Completo']
                            telefone = int(alunos.loc[x, "Whatsapp com DDD (somente números - sem espaço)"])

                            mensagem = "Olá *{0}.* Nós somos da AMTECH - Agência Maringá de Tecnologia e Inovação. entramos em contato porque você demonstrou interesse em cursos de tecnologia preenchendo um formulário.📋\n\nNós iremos iniciar em parceria com o *{1}*, o curso:  \n 🌟*{2}*.🌟 \n\nTodos podem participar desde que sejam maior de *{3}* anos e tenham a escolaridade mínima 5º ano do Ensino Fundamental.🎓\n\n🎯 Duração do curso: *{4}*\n 📆 Será no período *{5}*\n 🕒 Horário: *{6}* \n\n📩 Se você está interessado, responda esta mensagem para receber o formulário de inscrição.\n\n*⚠️ Atenção: As vagas são limitadas! Responda o mais rápido possível! 🏃‍♂️💨 📢*\n\n*📍Local: Acesso 1 | Piso Superior Terminal Urbano - Av. Tamandaré, 600 - Zona 01, Maringá -*\n\n*🏫 MO curso é PRESENCIAL E 100% GRATUITO! 🎉*".format(nome, parceiro, curso_de_envio, idademin, data_de_duracao, periodo, horario_do_curso)
                            # Criar links personalizados do whatsapp e enviar mensagens para cada cliente
                            link_mensagem_whatsapp = f'https://web.whatsapp.com/send/?phone={telefone}&text={quote(mensagem)}'
                            webbrowser.open(link_mensagem_whatsapp)
                            sleep(5)
                            # Automação para enviar a mensagem
                            sleep(5)
                            pyautogui.press('enter')
                            sleep(5)
                            pyautogui.hotkey('ctrl', 'w')
                else:
                    #olha cada item da lista criada, verifica se o curso de envio está dentro da lista, caso esteja, a planilha pega o nome e telefone do aluno, e faz o envio da mensagem.
                    for curso in lista_cursos:
                        if curso.upper() == curso_de_envio.upper():
                            # Ler planilha e guardar informações nome e telefone
                            nome = alunos.loc[x, 'Nome Completo']
                            telefone = int(alunos.loc[x, 'Whatsapp com DDD (somente números - sem espaço)'])

                            mensagem =  "Olá *{0}.* Nós somos da AMTECH - Agência Maringá de Tecnologia e Inovação. entramos em contato porque você demonstrou interesse em cursos de tecnologia preenchendo um formulário.📋\n\nNós iremos iniciar em parceria com o *{1}*, o curso:  \n 🌟*{2}*.🌟 \n\nTodos podem participar desde que sejam maior de *{3}* anos e tenham a escolaridade mínima 5º ano do Ensino Fundamental.🎓\n\n🎯 Duração do curso: *{4}*\n 📆 Será no período *{5}*\n 🕒 Horário: *{6}* \n\n📩 Se você está interessado, responda esta mensagem para receber o formulário de inscrição.\n\n*⚠️ Atenção: As vagas são limitadas! Responda o mais rápido possível! 🏃‍♂️💨 📢*\n\n*📍Local: Acesso 1 | Piso Superior Terminal Urbano - Av. Tamandaré, 600 - Zona 01, Maringá -*\n\n*🏫 MO curso é PRESENCIAL E 100% GRATUITO! 🎉*".format(nome, parceiro, curso_de_envio, idademin, data_de_duracao, periodo, horario_do_curso)

                            # Criar links personalizados do whatsapp e enviar mensagens para cada cliente
                            link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
                            webbrowser.open(link_mensagem_whatsapp)
                            sleep(5)
                            # Automação para enviar a mensagem
                            sleep(5)
                            pyautogui.press('enter')
                            sleep(5)
                            pyautogui.hotkey('ctrl', 'w')
                            
                       
        print("Todas as linhas foram lidas!")
        self.running = False
       

# Developed by: Lucas Ferrari Soares
# Contact: lucasferrarisoares@gmail.com
# DisparadordeMensagem 2.0    

# Criar a GUI
root = tk.Tk()
gui = CourseOfferGUI(root)
root.mainloop()
