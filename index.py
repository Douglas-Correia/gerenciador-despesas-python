from tkinter import *
from tkinter import messagebox
from tkinter import ttk
from tkinter import tix
from tkcalendar import Calendar, DateEntry
import pyodbc
from time import sleep
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, A4
from reportlab.pdfbase import pdfmetrics
from reportlab.platypus import SimpleDocTemplate, Image
import webbrowser


class SystemExpenditure():  # Expenditure = Despesas
    def __init__(self):
        self.Conexão_Banco()
        self.Tela_Frame()
        self.Lista_Frame()
        self.Campos_Label()
        self.Campos_Entry()
        self.ComboBox_Tipo()
        self.Variaveis_Banco()
        self.Campos_Button()
        self.Campos_Config()
        self.Limpa_Entrys()
        self.Add_Despesa()
        self.Select_Despesa()
        self.Config_Janela()

    def Conexão_Banco(self):
        self.banco_conexão = (
            "Driver={SQL Server};"
            "Server=DESKTOP-C3FJB6I\SQLEXPRESS;"
            "Database=BancoDespesa;"
        )
        self.conexão = pyodbc.connect(self.banco_conexão)
        self.cursor = self.conexão.cursor()
        print("conexão bem sucedida")

    def Tela_Frame(self):
        self.janela = tix.Tk()
        self.frame_title = Frame(self.janela, width=700, height=75, bg="#10454F", relief="flat")
        self.frame_title.grid(row=0, column=1, pady=1, padx=0, sticky=NSEW)

        self.frame_title_img = Frame(self.janela, width=100, height=74, bg="#10454F", relief="flat")
        self.frame_title_img.grid(row=0, column=0, pady=1, padx=0, sticky=NSEW)

        self.frame_menu = Frame(self.janela, width=100, height=424, bg="#506266", relief="flat")
        self.frame_menu.place(x=0, y=76.5)

        self.frame_conteudo = Frame(self.janela, width=599, height=423, bg="white", relief="flat")
        self.frame_conteudo.place(x=100.4, y=76.5)

    def Lista_Frame(self):
        self.listaDesp = ttk.Treeview(self.frame_conteudo, height=7, column=("col1", "col2", "col3", "col4"), show="headings")
        self.listaDesp.heading("#0", text="")
        self.listaDesp.heading("#1", text="Tipo")
        self.listaDesp.heading("#2", text="Nome")
        self.listaDesp.heading("#3", text="Preço")
        self.listaDesp.heading("#4", text="Data")

        # TAMANHO
        self.listaDesp.column("#0", width=1)
        self.listaDesp.column("#1", width=150)
        self.listaDesp.column("#2", width=200)
        self.listaDesp.column("#3", width=120)
        self.listaDesp.column("#4", width=120)
        # POSIÇÃO
        self.listaDesp.place(x=5, y=255)

        self.scroolLista = Scrollbar(self.frame_conteudo, orient="vertical")
        self.listaDesp.configure(yscroll=self.scroolLista.set)
        self.scroolLista.place(x=580, y=280, height=140)
        self.listaDesp.bind("<Double-1>", self.OndoubleClick)

    def Campos_Label(self):
        # LABEL
        self.lb_title = Label(self.frame_title, text="Gerenciador de Despesas", font="Ivy 15 bold", bg="#10454F",
                              fg="white", anchor=CENTER)

        self.logo = PhotoImage(file="image/image-desp80px.png")
        self.logo = self.logo.subsample(1, 1)
        self.figura1 = Label(self.frame_title_img, image=self.logo, bg="#10454F").grid(row=0, column=0)

        self.lb_tipo = Label(self.frame_conteudo, text="Tipo de despesa", bg="white", fg="black", font="Ivy 10 bold")
        self.lb_nome = Label(self.frame_conteudo, text="Nome da despesa", bg="white", fg="black", font="Ivy 10 bold")
        self.lb_preço = Label(self.frame_conteudo, text="Preço da despesa", bg="white", fg="black", font="Ivy 10 bold")
        self.lb_data = Label(self.frame_conteudo, text="Data da despesa", bg="white", fg="black", font="Ivy 10 bold")

    def Campos_Entry(self):
        self.ent_nome = Entry(self.frame_conteudo, width=40, relief="raised", bg="white", fg="black")
        self.ent_preço = Entry(self.frame_conteudo, width=15, relief="raised", bg="white", fg="black")
        self.ent_data = Entry(self.frame_conteudo, width=20, relief="raised", bg="white", fg="black")

        texto_balao_data = "Clique no botão Data para selecionar"
        self.balao_data = tix.Balloon(self.frame_conteudo)
        self.balao_data.bind_widget(self.ent_data, balloonmsg=texto_balao_data)

    def ComboBox_Tipo(self):
        listaTipos = ["Alimentação", "Veiculo", "Funcionário", "Casa", "Outros"]

        self.Combo_Tipos = ttk.Combobox(self.frame_conteudo, values=listaTipos)
        self.Combo_Tipos.set("Alimentação")

    def Campos_Button(self):
        self.bt_menu_lança = Button(self.frame_menu, text="Lançamentos", bd=2, width=13, height=3, bg="grey",
                                    fg="white", relief=RAISED, overrelief=RIDGE, command=self.Chamada_bt_menu)

        self.bt_menu_rela = Button(self.frame_menu, text="Relatórios", bd=2, width=13, height=3, bg="grey", fg="white",
                                   relief=RAISED, overrelief=RIDGE, command=self.Class_Relatório)

        self.bt_sair = Button(self.frame_menu, text="Sair", bd=2, width=13, height=3, bg="grey", fg="white",
                              relief=RAISED, overrelief=RIDGE, command=self.Sair_Sistema)

        texto_balao_button = "Para gerar o PDF, de 2 cliques na despesa desejada"
        self.balao_button = tix.Balloon(self.frame_menu)
        self.balao_button.bind_widget(self.bt_menu_rela, balloonmsg=texto_balao_button)

        self.bt_gravar = Button(self.frame_conteudo, text="Adicionar", bd=2, width=12, height=1, bg="#10454F", fg="white",
                                relief=RAISED, overrelief=RIDGE, command=self.Add_Despesa)

        self.bt_printDate = Button(self.frame_conteudo, text="Data", bd=2, width=7, height=1, bg="#10454F", fg="white",
                                   relief=RAISED, overrelief=RIDGE, command=self.Calendario)

        self.bt_deletar = Button(self.frame_conteudo, text="Apagar", bd=2, width=8, height=1, bg="#10454F", fg="white",
                                 relief=RAISED, overrelief=RIDGE, command=self.Excluir_dados)

        self.bt_limpar = Button(self.frame_conteudo, text="Limpar", bd=2, width=8, height=1, bg="#10454F", fg="white",
                                relief=RAISED, overrelief=RIDGE, command=self.Limpa_Entrys)

    def Calendario(self):
        self.calendario1 = Calendar(self.frame_conteudo, fg="grey75", bg="blue", font=("times", "9", "bold"),
                                    locale="pt_br")
        self.calendario1.place(x=377, y=70)

        self.bt_calendario = Button(self.frame_conteudo, text="Inserir data", bd=2, width=10, height=1, bg="#10454F",
                                    fg="white", relief=RAISED, overrelief=RIDGE, command=self.Print_calendario)
        self.bt_calendario.place(x=299, y=229)

    def Print_calendario(self):
        dataIni = self.calendario1.get_date()
        self.calendario1.destroy()
        self.ent_data.delete(0, END)
        self.ent_data.insert(END, dataIni)
        self.bt_calendario.destroy()

    def Campos_Config(self):
        # CONFIG LABEL
        self.lb_title.place(x=176, y=25)
        self.lb_tipo.place(x=16, y=10)
        self.lb_nome.place(x=16, y=60)
        self.lb_preço.place(x=310, y=10)
        self.lb_data.place(x=310, y=60)

        # CONFIG ENTRY
        # self.ent_tipo.place(x=20, y=30)
        self.ent_nome.place(x=20, y=80)
        self.ent_preço.place(x=314, y=30)
        self.ent_data.place(x=314, y=80)
        self.Combo_Tipos.place(x=20, y=35)

        # CONFIG BUTTON
        self.bt_menu_lança.place(x=0, y=0)
        self.bt_menu_rela.place(x=0, y=55)
        self.bt_sair.place(x=0, y=110)
        self.bt_gravar.place(x=20, y=120)
        self.bt_printDate.place(x=310, y=120)
        self.bt_deletar.place(x=380, y=120)
        self.bt_limpar.place(x=457, y=120)

    def Variaveis_Banco(self):
        self.combo_bd = self.Combo_Tipos.get()
        self.nome_bd = self.ent_nome.get()
        self.preço_bd = self.ent_preço.get()
        self.data_bd = self.ent_data.get()

    def Limpa_Entrys(self):
        self.Combo_Tipos.delete(0, END)
        self.ent_nome.delete(0, END)
        self.ent_preço.delete(0, END)
        self.ent_data.delete(0, END)

    def Add_Despesa(self):
        self.Variaveis_Banco()
        try:
            self.command_add = f"""INSERT INTO Despesas (Tipo, Nome, Preço, Datta) VALUES ('{self.combo_bd}', '{self.nome_bd}', {self.preço_bd}, '{self.data_bd}')"""
            self.cursor.execute(self.command_add)
            self.cursor.commit()
            messagebox.showinfo(title="Informações", message="Despesas gravadas")
            self.Limpa_Entrys()
            self.Select_Despesa()
        except:
            pass

    def Select_Despesa(self):
        self.listaDesp.delete(*self.listaDesp.get_children())
        lista = self.cursor.execute("""SELECT Tipo, Nome, Preço, Datta FROM Despesas ORDER BY Datta ASC;""")
        for i, (Tipo, Nome, Preço, Datta) in enumerate(lista, start=1):
            self.listaDesp.insert("", END, values=(Tipo, Nome, Preço, Datta))

    def OndoubleClick(self, event):
        self.Limpa_Entrys()
        self.listaDesp.selection()
        for n in self.listaDesp.selection():
            col1, col2, col3, col4 = self.listaDesp.item(n, 'values')
            self.Combo_Tipos.insert(END, col1)
            self.ent_nome.insert(END, col2)
            self.ent_preço.insert(END, col3)
            self.ent_data.insert(END, col4)

    def Excluir_dados(self):
        self.Variaveis_Banco()
        # self.cursor.execute("""DELETE FROM Despesas WHERE Nome = ?""", (self.nome_bd))
        deletar = f"""DELETE FROM Despesas WHERE Tipo = '{self.combo_bd}'"""
        self.cursor.execute(deletar)
        self.cursor.commit()
        self.Limpa_Entrys()
        self.Select_Despesa()

    def Chamada_bt_menu(self):
        if self.bt_menu_lança:
            self.janela.destroy()
            self.__init__()

    def Class_Relatório(self):
        sleep(0.5)
        messagebox.showinfo(title="Info Relatório", message="Relatório gerado em PDF")
        sleep(1)
        self.Gerar_PDF()

    def WebBrowser_PDF(self):
        webbrowser.open("Relatórios.pdf")

    def Gerar_PDF(self):
        self.Variaveis_Banco()
        self.canva = canvas.Canvas("Relatórios.pdf")

        self.canva.setFont("Helvetica-Bold", 24)
        self.canva.drawString(150, 790, "Relatórios de Despesas")

        self.canva.setFont("Helvetica-Bold", 15)
        self.canva.drawString(50, 700, "Tipo da Despesa: ")
        self.canva.drawString(50, 670, "Nome da Despesa: ")
        self.canva.drawString(50, 640, "Preço da Despesa: ")
        self.canva.drawString(50, 610, "Data da Despesa: ")

        self.canva.setFont("Helvetica", 12)
        self.canva.drawString(180, 700, self.combo_bd)
        self.canva.drawString(188, 670, self.nome_bd)
        self.canva.drawString(187, 640, self.preço_bd)
        self.canva.drawString(180, 610, self.data_bd)

        self.canva.rect(20, 550, 550, 5, fill=True, stroke=False)

        self.canva.showPage()
        self.canva.save()
        self.WebBrowser_PDF()

    def Sair_Sistema(self):
        quit()

    def Config_Janela(self):
        self.janela.geometry("700x500+300+90")
        self.janela.resizable(True, True)
        self.janela.maxsize(width=700, height=500)
        self.janela.minsize(width=400, height=300)
        self.janela.config(background="white")
        self.janela.title("Gerenciador de Despesas")
        self.janela.mainloop()

SystemExpenditure()