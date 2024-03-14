import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import datetime
from tkinter import messagebox

def enviar_email(remetente, senha, destinatarios, data_hora_envio, mensagem=None, arquivos=None):
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
        if arquivos:
            for arquivo_relatorio in arquivos:
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