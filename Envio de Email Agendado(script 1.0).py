import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkcalendar import Calendar, DateEntry
import datetime

class EmailSenderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Envio de Email Agendado")
        self.root.configure(bg="#212529")

        # Estilos
        self.entry_style = {'bg': '#343a40', 'fg': 'white', 'highlightbackground': '#212529', 'highlightcolor': '#007bff', 'insertbackground': 'white', 'borderwidth': 2, 'relief': 'solid'}
        self.button_style = {'bg': '#007bff', 'fg': 'white', 'activebackground': '#0056b3', 'activeforeground': 'white', 'borderwidth': 0, 'highlightthickness': 0, 'padx': 10, 'pady': 5, 'cursor': 'hand2'}
        self.label_style = {'bg': '#212529', 'fg': 'white'}

        # Configurações do e-mail
        self.remetente = tk.StringVar()
        self.remetente_label = tk.Label(root, text="Remetente:", **self.label_style)
        self.remetente_label.grid(row=0, column=0, padx=10, pady=5, sticky='w')
        self.remetente_entry = tk.Entry(root, textvariable=self.remetente, width=50, **self.entry_style)
        self.remetente_entry.grid(row=0, column=1, columnspan=2, padx=10, pady=5)

        self.senha = tk.StringVar()
        self.senha_label = tk.Label(root, text="Senha:", **self.label_style)
        self.senha_label.grid(row=1, column=0, padx=10, pady=5, sticky='w')
        self.senha_entry = tk.Entry(root, textvariable=self.senha, show="*", width=50, **self.entry_style)
        self.senha_entry.grid(row=1, column=1, columnspan=2, padx=10, pady=5)

        self.destinatarios = tk.StringVar()
        self.destinatarios_label = tk.Label(root, text="Destinatários (separados por vírgula):", **self.label_style)
        self.destinatarios_label.grid(row=2, column=0, padx=10, pady=5, sticky='w')
        self.destinatarios_entry = tk.Entry(root, textvariable=self.destinatarios, width=50, **self.entry_style)
        self.destinatarios_entry.grid(row=2, column=1, columnspan=2, padx=10, pady=5)

        # Seletor de Data
        self.data_label = tk.Label(root, text="Data Inicial:", **self.label_style)
        self.data_label.grid(row=3, column=0, padx=10, pady=5, sticky='w')
        self.data_entry = DateEntry(root, width=12, background='#343a40', foreground='white', borderwidth=2, locale='pt_BR')
        self.data_entry.grid(row=3, column=1, padx=10, pady=5)

        self.data_final_label = tk.Label(root, text="Data Final:", **self.label_style)
        self.data_final_label.grid(row=4, column=0, padx=10, pady=5, sticky='w')
        self.data_final_entry = DateEntry(root, width=12, background='#343a40', foreground='white', borderwidth=2, locale='pt_BR')
        self.data_final_entry.grid(row=4, column=1, padx=10, pady=5)

        # Seletor de Hora e Minuto
        self.hora_label = tk.Label(root, text="Hora:", **self.label_style)
        self.hora_label.grid(row=5, column=0, padx=10, pady=5, sticky='w')
        self.hora_entry = ttk.Combobox(root, values=[str(i).zfill(2) for i in range(24)], state="readonly", width=5)
        self.hora_entry.grid(row=5, column=1, padx=10, pady=5)

        self.minuto_label = tk.Label(root, text="Minuto:", **self.label_style)
        self.minuto_label.grid(row=5, column=2, padx=10, pady=5, sticky='w')
        self.minuto_entry = ttk.Combobox(root, values=[str(i).zfill(2) for i in range(60)], state="readonly", width=5)
        self.minuto_entry.grid(row=5, column=3, padx=10, pady=5)

        # Botão para selecionar os arquivos de relatório
        self.arquivos_relatorio = []
        self.arquivos_relatorio_label = tk.Label(root, text="Arquivos de Relatório:", **self.label_style)
        self.arquivos_relatorio_label.grid(row=6, column=0, padx=10, pady=5, sticky='w')
        self.selecionar_arquivos_button = tk.Button(root, text="Selecionar Arquivos", command=self.selecionar_arquivos, **self.button_style)
        self.selecionar_arquivos_button.grid(row=6, column=1, padx=10, pady=5)

        # Caixa de texto para mensagem
        self.mensagem_label = tk.Label(root, text="Mensagem de Texto:", **self.label_style)
        self.mensagem_label.grid(row=7, column=0, padx=10, pady=5, sticky='w')
        self.mensagem_text = tk.Text(root, width=50, height=5, **self.entry_style)
        self.mensagem_text.grid(row=7, column=1, columnspan=2, padx=10, pady=5)

        # Botão de envio de e-mail
        self.enviar_button = tk.Button(root, text="Agendar Envio", command=self.agendar_envio, **self.button_style)
        self.enviar_button.grid(row=8, column=1, columnspan=2, pady=10)

        # Botão para visualizar envios pendentes
        self.visualizar_button = tk.Button(root, text="Visualizar Envios Pendentes", command=self.visualizar_envios_pendentes, **self.button_style)
        self.visualizar_button.grid(row=9, column=1, columnspan=2, pady=10)

        # Lista de envios agendados
        self.envios_agendados = []

    def selecionar_arquivos(self):
        filenames = filedialog.askopenfilenames(initialdir="/", title="Selecionar Arquivos", filetypes=(("Arquivos de Texto", "*.txt"), ("Todos os arquivos", "*.*")))
        self.arquivos_relatorio = list(filenames)
        self.arquivos_relatorio_label.config(text=f"Arquivos de Relatório: {len(self.arquivos_relatorio)} arquivos selecionados")

    def agendar_envio(self):
        remetente = self.remetente.get()
        senha = self.senha.get()
        destinatarios = self.destinatarios.get()
        data_inicial = self.data_entry.get_date()
        data_final = self.data_final_entry.get_date()
        hora = self.hora_entry.get()
        minuto = self.minuto_entry.get()
        mensagem = self.mensagem_text.get("1.0", tk.END).strip()

        if not remetente or not senha or not destinatarios:
            messagebox.showerror("Erro", "Preencha todos os campos obrigatórios.")
            return

        if not self.arquivos_relatorio and not mensagem:
            messagebox.showerror("Erro", "Selecione pelo menos um arquivo de relatório ou insira uma mensagem de texto.")
            return

        data_hora_envio = datetime.datetime.combine(data_inicial, datetime.time(int(hora), int(minuto)))

        # Agendar o envio para cada dia no intervalo selecionado
        delta_dias = (data_final - data_inicial).days + 1
        for i in range(delta_dias):
            data_envio = data_inicial + datetime.timedelta(days=i)
            data_hora_envio = datetime.datetime.combine(data_envio, datetime.time(int(hora), int(minuto)))
            if mensagem and not self.arquivos_relatorio:
                self.enviar_email(remetente, senha, destinatarios, data_hora_envio, mensagem)
            elif self.arquivos_relatorio and not mensagem:
                self.enviar_email(remetente, senha, destinatarios, data_hora_envio, None, self.arquivos_relatorio)
            else:
                self.enviar_email(remetente, senha, destinatarios, data_hora_envio, mensagem, self.arquivos_relatorio)

    def enviar_email(self, remetente, senha, destinatarios, data_hora_envio, mensagem=None, arquivos=None):
        servidor = None
        try:
            # Criando o objeto de mensagem
            msg = MIMEMultipart()
            msg['From'] = remetente
            msg['To'] = destinatarios
            msg['Subject'] = 'Relatório Diário'

            # Corpo da mensagem
            if not mensagem:
                corpo = "Olá,\n\nSegue em anexo o relatório diário.\n\nAtenciosamente,\nSeu Nome"
            else:
                corpo = mensagem
            msg.attach(MIMEText(corpo, 'plain'))

            # Anexando os relatórios
            for arquivo_relatorio in self.arquivos_relatorio:
                with open(arquivo_relatorio, 'rb') as anexo:
                    parte_anexo = MIMEBase('application', 'octet-stream')
                    parte_anexo.set_payload(anexo.read())
                encoders.encode_base64(parte_anexo)
                parte_anexo.add_header('Content-Disposition', f"attachment; filename= {arquivo_relatorio}")
                msg.attach(parte_anexo)

            # Conectando ao servidor SMTP e enviando o e-mail
            servidor_smtp = 'smtp.gmail.com'
            porta_smtp = 587
            servidor = smtplib.SMTP(host=servidor_smtp, port=porta_smtp)
            servidor.starttls()
            servidor.login(remetente, senha)
            texto_email = msg.as_string()
            servidor.sendmail(remetente, destinatarios.split(','), texto_email)
            messagebox.showinfo("Sucesso", "E-mail agendado com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao enviar e-mail: {e}")
        finally:
            if servidor:
                servidor.quit()

    def visualizar_envios_pendentes(self):
        # Criar uma nova janela para exibir os envios pendentes
        self.janela_visualizacao = tk.Toplevel(self.root)
        self.janela_visualizacao.title("Envios Pendentes")
        self.janela_visualizacao.configure(bg="#212529")

        # Criar uma tabela para exibir os envios pendentes
        columns = ("Remetente", "Destinatários", "Data e Hora", "Mensagem", "Arquivos")
        self.lista_envios_pendentes = ttk.Treeview(self.janela_visualizacao, columns=columns, show="headings", selectmode="browse")
        for col in columns:
            self.lista_envios_pendentes.heading(col, text=col)
            self.lista_envios_pendentes.column(col, width=150, anchor="center")
        self.lista_envios_pendentes.pack(expand=True, fill="both")

        # Preencher a tabela com os envios pendentes
        self.atualizar_lista_envios_pendentes()

        # Adicionar um botão para cancelar envio
        self.cancelar_button = tk.Button(self.janela_visualizacao, text="Cancelar Envio", command=self.cancelar_envio_selecionado, **self.button_style)
        self.cancelar_button.pack(pady=10)


    def atualizar_lista_envios_pendentes(self):
        # Limpar a lista de envios pendentes antes de atualizar
        for envio in self.envios_agendados:
            self.lista_envios_pendentes.delete(envio)

        # Preencher a lista de envios pendentes com os detalhes dos envios agendados
        for i, envio in enumerate(self.envios_agendados):
            remetente, destinatarios, data_hora_envio, mensagem, arquivos = envio
            data_hora_str = data_hora_envio.strftime("%Y-%m-%d %H:%M")
            self.lista_envios_pendentes.insert("", "end", values=(remetente, destinatarios, data_hora_str, mensagem, arquivos))

    def cancelar_envio_selecionado(self):
        # Obter o envio selecionado na lista
        selected_item = self.lista_envios_pendentes.selection()
        if not selected_item:
            messagebox.showerror("Erro", "Nenhum envio selecionado.")
            return

        # Remover o envio selecionado da lista de envios agendados
        index = int(selected_item[0][1:]) - 1
        del self.envios_agendados[index]

        # Atualizar a lista de envios pendentes na interface
        self.atualizar_lista_envios_pendentes()

root = tk.Tk()
app = EmailSenderApp(root)
root.mainloop()