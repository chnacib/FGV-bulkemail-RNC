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
from PIL import Image, ImageTk
import time


host = ""
port = ""
login = ""
senha = ""
cc = ""


class App:
    def validate(self):
        login_entry  = e1.get()
        passwd_entry = e2.get()
        server = smtplib.SMTP(host, port)
        app.login_entry = login_entry
        app.passwd_entry = passwd_entry
        try:
            server.ehlo()
            server.starttls()
            server.login(login_entry, passwd_entry)
            server.quit()
            conect_label = Label(root,text="Login efetuado com sucesso.",fg="green")
            conect_label.place(x=160,y=370,width=300,height=30)
            root.update()
            time.sleep(5)
            conect_label.destroy()
            app.destroy_layer()
            app.menu()
        except:
            except_label = Label(root,text="Login ou senha incorretos.",fg="red")
            except_label.place(x=160,y=370,width=300,height=30)



    def destroy_layer(self):
        e1.destroy()
        e2.destroy()
        login_txt.destroy()
        senha_txt.destroy()
        login_btn.destroy()
        x.destroy()
        app.menu()


    def month_submit(self,month_selected):
        month_selected = app.variableget
        app.month_selected = str(month_selected)
        submit_text = tk.Label(root,text="Mês selecionado!" ,fg='green')
        submit_text.place(x=240,y=310,width=140,height=20)
    

    def importar_planilha(self):
        a = askopenfilename()  # Isto te permite selecionar um arquivo
        df = pd.read_excel(a)
        a_path = os.path.abspath(a)
        planilha_path = tk.Label(root, text=a_path,fg='green')
        planilha_path.place(x=160,y=135,width=300,height=20)
        app.df = df
    def importar_diretorio(self):
        b = filedialog.askdirectory()
        b_path = os.path.abspath(b)
        directory_path = tk.Label(root, text=b_path, fg='green')
        directory_path.place(x=120, y=195, width=400, height=20)
        app.b_path = str(b_path)
    def importar_texto(self):
        c = askopenfilename()
        c_path = os.path.abspath(c)
        txtfile = open(c,encoding='utf8')
        contents = txtfile.read()
        text_status = tk.Label(root,text=c_path,fg='green')
        text_status.place(x=160,y=255,width=300,height=20)
        app.contents = contents
    def send_email(self):
        server = smtplib.SMTP(host, port)
        server.ehlo()
        server.starttls()
        server.login(app.login_entry, app.passwd_entry)
        listaemail = app.df['E-mail']
        listaemail = list(listaemail)
        listanome = app.df['CT']
        listanome = list(listanome)
        cont = 1
        for email,nome in zip(listaemail,listanome):
            filepath = f"{app.b_path}\page-{cont}.docx"
            newpath = filepath.replace('\\','/')
            print(filepath)
            body = app.contents
            email_msg = MIMEMultipart()
            email_msg['From'] = app.login_entry
            email_msg['To'] = email
            email_msg['Cc'] = cc
            email_msg['Subject'] = f"Relatório de não conformidade - {app.month_selected} - {nome}"
            rcpt = []
            rcpt.append(email_msg['To'])
            rcpt.append(email_msg['Cc'])
            email_msg.attach(MIMEText(body, 'html'))
            cam_arquivo = newpath
            attchment = open(cam_arquivo,'rb')
            att = MIMEBase('application', 'octet-stream')
            att.set_payload(attchment.read())
            encoders.encode_base64(att)
            att.add_header('Content-Disposition', f'attachment', filename=f'page-{cont}.docx')
            attchment.close()
            email_msg.attach(att)
            server.sendmail(email_msg['From'], rcpt, email_msg.as_string())
            cont = cont + 1

        server.quit()
        send_status = tk.Label(root, text='Email enviado!', fg='green')
        send_status.place(x=160,y=370,width=300,height=30)
    def menu(self):
        #Botão importsheet
        planilha_btn = Button(root,text='Importar Planilha',command=app.importar_planilha)
        planilha_btn.place(x=240,y=100,width=140,height=30)
        #Choose directory
        directory_btn = Button(root,text='Importar Diretório',command=app.importar_diretorio)
        directory_btn.place(x=240,y=160,width=140,height=30)
        #Choose text
        text_btn = Button(root,text='Importar texto',command=app.importar_texto)
        text_btn.place(x=240,y=220,width=140,height=30)
        #Send
        send_btn = Button(root,text='Enviar',command=app.send_email)
        send_btn.place(x=240,y=335,width=140,height=30)
        #Month
        variable = StringVar(root)
        variable.set("Janeiro") # default value

        w = OptionMenu(root,variable,'Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro',command=app.month_submit)
        w.place(x=240,y=280,width=140,height=30)
        app.variableget = variable.get()


app = App()


month_opt = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro']






root = Tk()
root.title("FGV CertPessoas")
root.iconphoto(False, tk.PhotoImage(file='fgvlogo1.png'))
root.geometry('600x400')
#Titulo
imagelogo = tk.PhotoImage(file="projetoslogo2.png")
z = tk.Label(image=imagelogo)
z.grid(padx=110,pady=0)
imagelogo2 = tk.PhotoImage(file="userlogo.png")
x = tk.Label(image=imagelogo2)
x.grid(padx=40,pady=0)

#login
login_txt = tk.Label(root,text="Login:",fg='blue',font=60)
login_txt.place(x=160,y=250,width=60,height=30)
e1 = tk.Entry(root)
e1.place(x=220,y=250,width=200,height=30)
#senha
senha_txt = tk.Label(root,text="Senha:",fg='blue',font=60)
senha_txt.place(x=160,y=300,width=60,height=30)
e2 = tk.Entry(root,show="*")
e2.place(x=220,y=300,width=200,height=30)

login_btn = Button(root,text='Entrar',command=app.validate)
login_btn.place(x=240,y=340,width=140,height=30)





root.mainloop()