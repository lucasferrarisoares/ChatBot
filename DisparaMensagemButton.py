# Importar bibliotecas necessárias
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


# Criar uma classe para a GUI
class CourseOfferGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Disparador de Mensagem!")
        self.root.geometry("700x600")


        #TÍTULO
        app = self.root
        label = ttk.Label(app, text='Responda para começarmos!')
        label.pack(pady=35)
        label.config(font=('Arial', 20, 'bold'))
        style = Style(theme='superhero')

        # Recebe os dados de entrada.
        self.menu_course = ttk.Menubutton(root, text="Curso")
        self.course_selected = tk.StringVar()
        self.course = tk.Menu(self.menu_course, tearoff=0)
        self.course.add_command(label="REDES SOCIAIS: UM GUIA PRÁTICO  PARA ALAVANCAR SUAS VENDAS", command=lambda: self.Chosing_Course("REDES SOCIAIS: UM GUIA PRÁTICO  PARA ALAVANCAR SUAS VENDAS"))
        self.course.add_command(label="CRIAÇÃO DE MÍDIAS PARA REDES SOCIAIS", command=lambda: self.Chosing_Course("CRIAÇÃO DE MÍDIAS PARA REDES SOCIAIS"))
        self.course.add_command(label="EXCEL - INTERMEDIÁRIO", command=lambda: self.Chosing_Course("EXCEL - INTERMEDIÁRIO"))
        self.course.add_command(label="INFORMÁTICA BÁSICA (TERCEIRA IDADE)", command=lambda: self.Chosing_Course("INFORMÁTICA BÁSICA (TERCEIRA IDADE)"))
        self.course.add_command(label="INSTAGRAM PARA NEGÓCIOS", command=lambda: self.Chosing_Course("INSTAGRAM PARA NEGÓCIOS"))
        self.course.add_command(label="INTRODUÇÃO A LÓGICA DE PROGRAMAÇÃO", command=lambda: self.Chosing_Course("INTRODUÇÃO A LÓGICA DE PROGRAMAÇÃO"))
        self.course.add_command(label="INTRODUÇÃO A INFORMÁTICA (WINDOWS; WORD, EXCEL; POWERPOINT E INTERNET)", command=lambda: self.Chosing_Course("INTRODUÇÃO A INFORMÁTICA (WINDOWS; WORD, EXCEL; POWERPOINT E INTERNET)"))
        self.course.add_command(label="INTRODUÇÃO A LÓGICA DE PROGRAMAÇÃO COM ALGORÍTIMOS", command=lambda: self.Chosing_Course("INTRODUÇÃO A LÓGICA DE PROGRAMAÇÃO COM ALGORÍTIMOS"))
        self.course.add_command(label="PACOTE OFFICE", command=lambda: self.Chosing_Course("PACOTE OFFICE"))
        self.course.add_command(label="SKETCHUP", command=lambda: self.Chosing_Course("SKETCHUP"))
        self.menu_course.configure(menu=self.course)
        self.menu_course.pack(pady=5)

        self.menu_button = ttk.Menubutton(root, text="Instituição Parceira")
        self.partner_selected = tk.StringVar()
        self.parceiro = tk.Menu(self.menu_button, tearoff=0)
        self.parceiro.add_command(label="SENAC", command=lambda: self.Chosing_Partner("SENAC"))
        self.parceiro.add_command(label="SENAI", command=lambda: self.Chosing_Partner("SENAI"))
        self.parceiro.add_command(label="SEBRAE", command=lambda: self.Chosing_Partner("SEBRAE"))
        self.menu_button.configure(menu=self.parceiro)
        self.menu_button.pack(pady=5)

        self.menu_period = ttk.Menubutton(root, text="Período")
        self.period_selected = tk.StringVar()
        self.period = tk.Menu(self.menu_period, tearoff=0)
        self.period.add_command(label="MATUTINO", command=lambda: self.Chosing_Period("MATUTINO"))
        self.period.add_command(label="VESPERTINO", command=lambda: self.Chosing_Period("VESPERTINO"))
        self.period.add_command(label="NOTURNO", command=lambda: self.Chosing_Period("NOTURNO"))
        self.menu_period.configure(menu=self.period)
        self.menu_period.pack(pady=5)

        self.menu_group = ttk.Menubutton(root, text="Deseja enviar mensagem por grupos?")
        self.group_selected = tk.StringVar()
        self.group = tk.Menu(self.menu_period, tearoff=0)
        self.group.add_command(label="SIM", command=lambda: self.Chosing_Group("SIM"))
        self.group.add_command(label="NÃO", command=lambda: self.Chosing_Group("NÃO"))
        self.menu_group.configure(menu=self.group)
        self.menu_group.pack(pady=5)

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

        self.minrange_label = ttk.Label(root, text="De qual linha devo começar:")
        self.minrange_label.pack()
        self.minrange_entry = ttk.Entry(root, width=30)
        self.minrange_entry.pack()

        self.maxrange_label = ttk.Label(root, text="Até qual linha devo enviar:")
        self.maxrange_label.pack()
        self.maxrange_entry = ttk.Entry(root, width=30)
        self.maxrange_entry.pack()

        # Criar um botão para enviar mensagens
        self.send_button = ttk.Button(root, text="Enviar Mensagens", command=self.start_sending)
        self.send_button.pack(pady=10)

        # Criar um botão para cancelar o código
        self.cancel_button = ttk.Button(root, text="Interromper código", bootstyle=DANGER, command=self.interromper_codigo)
        self.cancel_button.pack(pady=10)

        

            

        self.running = False

    #funcoes suporte.
    def interromper_codigo(self):
        print("Envio encerrado")
        self.running = False

    def start_sending(self):
        if self.running:
            print("O envio de mensagens já está em andamento.")
            return
        self.running = True
        print('Iniciado o envio de mensagens!')
        t = threading.Thread(target=self.send_messages)
        t.start()
    
    def Chosing_Course(self, curso: str):
        self.menu_course.config(text=curso)
        self.course_selected.set(curso)

    def Chosing_Partner(self, parceiro: str):
        self.menu_button.config(text=parceiro)
        self.partner_selected.set(parceiro)

    def Chosing_Period(self, periodo: str):
        self.menu_period.config(text=periodo)
        self.period_selected.set(periodo)
    
    def Chosing_Group(self, grupo: str):
        self.menu_group.config(text=grupo)
        self.group_selected.set(grupo)

    def encontra_categoria(self, curso_de_envio: str) -> list:
        #DEFINE AS CATEGORIAS
        Dev=["INTRODUÇÃO A LÓGICA DE PROGRAMAÇÃO", "INTRODUÇÃO A LÓGICA DE PROGRAMAÇÃO COM ALGORÍTIMOS", "Montagem e Manutenção de Computadores e Redes", "Programação, Redes e programação", 
             "SQL Linguagem de programação", "Inteligência artificial", "Cloud", "DBA", "Big Data", "Desenvolvimento web", "Programação Avançada", "Introdução a Programação", "Lógica de Programação"]

        Marketing=["REDES SOCIAIS: UM GUIA PRÁTICO PARA ALAVANCAR SUAS VENDAS", "CRIAÇÃO DE MÍDIAS PARA REDES SOCIAIS", "INSTAGRAM PARA NEGÓCIOS", "Marketing Digital"]      

        Design=["CRIAÇÃO DE MÍDIAS PARA REDES SOCIAIS", "SKETCHUP", "Design", "SKETCHUP - DESENHOS DE AMBIENTES E OBJETOS EM 3D - NOITE", "SketchUp 3D"]

        Pacote_Office=["EXCEL - INTERMEDIÁRIO", "EXCEL - AVANÇADO", "PACOTE OFFICE", "Edição de Texto", "Planilha Eletrônica"]

        Basicos=["INFORMÁTICA BÁSICA (TERCEIRA IDADE)", "INTRODUÇÃO A INFORMÁTICA (WINDOWS; WORD, EXCEL; POWERPOINT E INTERNET)", "Informática básica", "Informática básica (terceira idade)", 
                 "Informática básica a noite", "Informática básica (jovens)", "INFORMÁTICA BÁSICA (TERCEIRA IDADE) - MANHÃ", "Informatica basica a noite", "Introdução de Informática - TARDE - 22/01 à 21/02/24", 
                 "Introdução de Informática"]
        lista_categoria = [Dev, Marketing, Design, Pacote_Office, Basicos]

        Correto = None
        while Correto == None:
            for Categorias in lista_categoria:
                if curso_de_envio in Categorias:
                    Correto = Categorias
        return Correto


    def send_messages(self):
        # Salva as entradas obtidas na interface.
        curso_de_envio = self.course_selected.get()
        parceiro = self.partner_selected.get()
        periodo = self.period_selected.get()
        horario_do_curso = self.schedule_entry.get()
        data_de_duracao = self.duration_entry.get()
        linhamin = int(self.minrange_entry.get())
        linhamax = int(self.maxrange_entry.get())
        idademin = int(self.minage_entry.get())
        por_grupo = self.group_selected.get()
    

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

            cursos = alunos.loc[x, "Dentre as opções qual curso gostaria de fazer?"]
            lista_cursos = cursos.split(sep=', ')

            if por_grupo == "SIM":
                categoria = self.encontra_categoria( curso_de_envio)
                for curso in lista_cursos:
                    if curso in categoria:
                        # Ler planilha e guardar informações nome e telefone
                        nome = alunos.loc[x, 'Nome Completo']
                        telefone = alunos.loc[x, 'Whatsapp com DDD (somente números - sem espaço)']

                        mensagem = "Olá {0}. Nós somos da AMTECH - Agência Maringá de Tecnologia e Inovação. estamos entrando em contato pois você respondeu um formulário de interesse em cursos na área de tecnologia.\nNós iremos iniciar em parceria com o {1}, o curso {2}.\nTodos podem participar desde que sejam maior de {3} anos e tenham a escolaridade mínima 5º ano do Ensino Fundamental.\nO curso tem duração do dia {4} e será no período {5} das {6}\nAqueles que tiverem interesse, favor respondam essa mensagem, que iremos enviar o formulário para preenchimento dos dados\nATENÇÃO!\nPois as vagas são LIMITADAS!!".format(nome, parceiro, curso_de_envio, idademin, data_de_duracao, periodo, horario_do_curso)

                        # Criar links personalizados do whatsapp e enviar mensagens para cada cliente
                        link_mensagem_whatsapp = f'https://api.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
                        webbrowser.open(link_mensagem_whatsapp)
                        sleep(5)
                        break
            else:
                #olha cada item da lista criada, verifica se o curso de envio está dentro da lista, caso esteja, a planilha pega o nome e telefone do aluno, e faz o envio da mensagem.
                for curso in lista_cursos:
                    if curso.upper() == curso_de_envio.upper():
                        # Ler planilha e guardar informações nome e telefone
                        nome = alunos.loc[x, 'Nome Completo']
                        telefone = alunos.loc[x, 'Whatsapp com DDD (somente números - sem espaço)']

                        mensagem = "Olá {0}. Nós somos da AMTECH - Agência Maringá de Tecnologia e Inovação. estamos entrando em contato pois você respondeu um formulário de interesse em cursos na área de tecnologia.\nNós iremos iniciar em parceria com o {1}, o curso {2}.\nTodos podem participar desde que sejam maior de {3} anos e tenham a escolaridade mínima 5º ano do Ensino Fundamental.\nO curso tem duração do dia {4} e será no período {5} das {6}\nAqueles que tiverem interesse, favor respondam essa mensagem, que iremos enviar o formulário para preenchimento dos dados\nATENÇÃO!\nPois as vagas são LIMITADAS!!".format(nome, parceiro, curso_de_envio, idademin, data_de_duracao, periodo, horario_do_curso)

                        # Criar links personalizados do whatsapp e enviar mensagens para cada cliente
                        link_mensagem_whatsapp = f'https://api.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
                        webbrowser.open(link_mensagem_whatsapp)
                        sleep(5)

        print("Todas as linhas foram lidas!")
        self.running = False
       
# Criar a GUI
root = tk.Tk()
gui = CourseOfferGUI(root)
root.mainloop()

