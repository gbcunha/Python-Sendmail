from cgitb import html
from tkinter import *
from tkinter import ttk
import pandas as pd
from tkinter.filedialog import askopenfilename
import os
import tkinter as tk
from tkinter import Tk, filedialog
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import smtplib 




a = askopenfilename()  # Isto te permite selecionar um arquivo
df = pd.read_excel(a)
listaemail = df['Email']
listaemail = list(listaemail)
listanome = df['CT']
listanome = list(listanome)
listacont = df['Punctuation']
listacont = list(listacont)
listacont2 = df['Position']
listacont2 = list(listacont2)

print (listaemail)

host = "smtp.office365.com"
port = "587"
login = "E-MAIL" # Neste campo você deve inserir o e-mail que será enviado.
senha = "PASSWORD" # Neste campo você deve inserir a senha do e-mail que será enviado.
cc = "****" # Neste campo você deve inserir o que será copiado, caso necessário.

server = smtplib.SMTP(host,port)
server.ehlo()
server.starttls()
server.login(login,senha)


cont = 1
cont2 = 1
for email,nome,cont,cont2 in zip(listaemail,listanome,listacont,listacont2):
    # Corpo do e-mail 
    body = f""" <p style="font-family:calibri"> Prezado(a) colaborador(a); </p> 
                <p>                            </p> 
                <p style="font-family:calibri">Segue abaixo a posição do seu Centro de Testes no Ranking gerado com base na aplicação dos exames de certificação durante o mês de Julho de 2022</p>
                <p>                            </p> 
               <p style="font-family:calibri"><b>Centro de Testes: {nome}</b></p>  
                <p>                            </p> 
                <p style="font-family:calibri"><b>Pontuação: {cont}</b></p> 
                <p>                            </p> 
                <p style="font-family:calibri"><b>Posição: {cont2}º</b></p>
                <p style="font-family:calibri">Qualquer dúvida estamos à disposição.</p>
                <p>                            </p> 
                <p style="font-family:calibri">Grato.</p>
                <p>                            </p>
                <p style="font-family:calibri">At.te,</p>
                <p>                            </p> 
                <p style="font-family:calibri">
                    <b>Gabrie Cunha</b>
                        <br><i>Certificação de Pessoas</i></br>
                            <br><b>Tel.:(21) 3799 4645</b></br></p>

                <p style="font-family:calibri">gabriel.cunha@fgv.br</p> 
                """

    email_msg = MIMEMultipart()
    email_msg['From'] = login
    email_msg['To'] = email
    email_msg['Cc'] = cc
    email_msg['Subject'] = f"Ranking FGV Projetos - Setembro - {nome}" # Neste campo você deve inserir o assunto do e-mail.
    rcpt = []
    rcpt.append(email_msg['To'])
    rcpt.append(email_msg['Cc'])
    email_msg.attach(MIMEText(body,'html'))
    server.sendmail(email_msg['From'],rcpt,email_msg.as_string())
    cont = cont + 1
    cont2 = cont2 + 1

server.quit()

print('E-mails enviados...')
