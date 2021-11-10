from tkinter import *
from tkinter.font import Font
import tkinter.messagebox
import smtplib
import pandas as pd

def send_message():
    your_name="Example"
    your_email="012example@gmail.com"
    your_password="Example@012"
    
#creating a secure connection with gmail.
    server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
    server.ehlo()
    server.login(your_email, your_password)
    subject_info = sub.get()
    email_body_info = body.get()

    email_list=pd.read_excel("Book1.xlsx")
    print(email_list.head())
    all_names=email_list['name']
    all_emails=email_list['email']

    for idx in range(len(all_emails)):
      name=all_names[idx]
      email=all_emails[idx]
      
      msg=f'Subject: {subject_info}\n\n\nHello, {name} \n{email_body_info} \nThank You {your_name}'
      try:
            server.sendmail(your_email,[email], msg)
            print("Email to {} successfully sent\n\n".format(email))


      except Exception as e:
            print("Error, could not send.")

    email_body_entry.delete(0,END)
    e1.delete(0,END)
    tkinter.messagebox.showinfo("DONE","Mails Sent Successfully!")
    

window=Tk()
window.title("AUTOMATIC MAILING SYSTEM")
window.geometry()

top_frame=Frame(window).pack()
bottom_frame=Frame(window).pack(side='bottom')

label=Label(top_frame, text="PYTHON E-MAIL SENDING APP", bg='black', fg='white',font=('Courier New',25), width='500',height='3').pack(padx=3)
canvas = Canvas(top_frame, width = 150, height = 115)      
canvas.pack()      
pic=PhotoImage(file="GUI.gif")
canvas.create_image(20,20, anchor=NW, image=pic) 


label1=Label(bottom_frame, text="SUBJECT FOR THE MAIL", font=("Times New Roman",15)).pack(pady=10)
sub=StringVar()
e1=Entry(bottom_frame,textvariable=sub,width=30, highlightthickness=1)
e1.pack()

label2=Label(bottom_frame, text="BODY FOR THE MAIL", font=("Times New Roman",15)).pack(pady=10)

body=StringVar()
email_body_entry=Entry(bottom_frame,textvariable=body,width=70, highlightthickness=1)
email_body_entry.pack()
button=Button(bottom_frame,text='SEND MESSAGE', font=("Arial Hebrew",10),bg='grey',height='2',width='10', command=send_message).pack(ipadx=20,pady=40)

window.mainloop()