from tkinter import *
import tkinter as tk
import tkinter.messagebox as mbox
from tkinter import filedialog
import os
from xlrd import open_workbook
import urllib.request,urllib.parse
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import time
import re, uuid
def validate():
    mac=':'.join(re.findall('..', '%012x' % uuid.getnode()))
    mf=open("m.act.txt","r")
    if mac in mf.read():
        master()
    else:
        mbox.showinfo("Failed","This Computer Not A Member!")
def master():
 global sc1
 sc1=tk.Tk()
 sc1.title('LOGIN')
 sc1.minsize(340,180)
 sc1.maxsize(340,180)
 f=tk.Frame(sc1)
 f.pack(side=tk.TOP,fill=tk.BOTH)
 l=tk.Label(f,text="LOGIN",bg="brown",fg="white",font=('bold',27))
 l.pack(fill=tk.X)
 f4=tk.Frame(sc1)
 f4.pack(fill=tk.BOTH)
 global code
 code=tk.Entry(f4,bg="burlywood",width=25,font=(20),show="*")
 code.pack(side=tk.LEFT,fill=tk.X,pady=45,padx=10)
 b=tk.Button(f4,text="LOGIN",command=login,bg="blue",fg="white")
 sc1.after(1, lambda:code.focus_force() )
 sc1.bind('<Return>',lambda event:login())
 b.pack(side=tk.LEFT,fill=tk.X,pady=15,padx=10)
 sc1.mainloop()
 return
def login():
 global c
 c=code.get()
 f=open("run.txt","r")
 passw=f.read()
 f.close()
 if(c==passw):
  mbox.showinfo("Success","Login Success!")
  sc1.destroy()
  main()
 else:
  mbox.showinfo("Failed","Incorect Password!")
 return
def main():
 global reset,sc
 reset=0
 sc=tk.Tk()
 sc.title('HOME')
 sc.minsize(962,500)
 sc.maxsize(962,500)
 sb=tk.Scrollbar(sc)
 sb.pack(side=tk.RIGHT,fill=tk.Y)
 l=tk.Label(sc,text="PayOuT",bg="brown",fg="white",font=('bold'))
 l.pack(fill=tk.X)
 menu_bar = tk.Menu(sc)
 global account 
 account= tk.Menu(menu_bar, tearoff=False)
 account.add_command(label="Login", command=loginf)
 account.add_command(label="Logout", command=logoutt)
 account.add_command(label="Change Password", command=chpass)
 account.entryconfig(1,state=DISABLED)
 menu_bar.add_cascade(label="Account", menu=account)
 sc.config(menu=menu_bar)
 f2b=tk.Frame(sc)
 f2b.pack(fill=tk.BOTH)
 gpflb=tk.Label(f2b,text="MONTH                      :",font=('bold'))
 gpflb.pack(side=tk.LEFT,fill=tk.X,pady=15,padx=10)
 global month
 month=tk.Entry(f2b,bg="burlywood",width=20,font=(20))
 month.pack(side=tk.LEFT,fill=tk.X,pady=15,padx=10)
 f4=tk.Frame(sc)
 f4.pack(fill=tk.BOTH,pady=10)
 l.config(font=('Arial',10))
 usrl=tk.Label(f4,text="HTML FILE NAME    :",font=('bold'))
 usrl.pack(side=tk.LEFT,fill=tk.X,pady=5,padx=10)
 global file
 file=tk.Entry(f4,bg="burlywood",width=62,font=(20))
 file.pack(side=tk.LEFT,fill=tk.X,pady=5,padx=10)
 file.config(state='readonly')
 browseb=tk.Button(f4,text="BROWSE",command=lambda:browse(1),bg="blue",fg="white")
 browseb.pack(side=tk.LEFT,fill=tk.X,pady=5,padx=10)
 f5=tk.Frame(sc)
 f5.pack(fill=tk.BOTH,pady=10)
 usrl1=tk.Label(f5,text="HTML FILE NAME 2 :",font=('bold'))
 usrl1.pack(side=tk.LEFT,fill=tk.X,pady=5,padx=10)
 global file1
 file1=tk.Entry(f5,bg="burlywood",width=62,font=(20))
 file1.pack(side=tk.LEFT,fill=tk.X,pady=5,padx=10)
 file1.config(state='readonly')
 browseb1=tk.Button(f5,text="BROWSE",command=lambda:browse(2),bg="blue",fg="white")
 browseb1.pack(side=tk.LEFT,fill=tk.X,pady=5,padx=10)
 f4b=tk.Frame(sc)
 f4b.pack(fill=tk.BOTH,pady=10)
 usrlb=tk.Label(f4b,text="EXCEL FILE NAME :",font=('bold'))
 usrlb.pack(side=tk.LEFT,fill=tk.X,pady=5,padx=10)
 global fileb
 fileb=tk.Entry(f4b,bg="burlywood",width=62,font=(20))
 fileb.pack(side=tk.LEFT,fill=tk.X,pady=5,padx=10)
 fileb.config(state='readonly')
 browsebb=tk.Button(f4b,text="BROWSE",command=browsefb,bg="blue",fg="white")
 browsebb.pack(side=tk.LEFT,fill=tk.X,pady=5,padx=10)
 f2=tk.Frame(sc)
 f2.pack(fill=tk.BOTH)
 gpfl=tk.Label(f2,text="Employee_ID           :",font=('bold'))
 gpfl.pack(side=tk.LEFT,fill=tk.X,pady=15,padx=10)
 global gpfno
 gpfno=tk.Entry(f2,bg="burlywood",width=28,font=(20))
 gpfno.pack(side=tk.LEFT,fill=tk.X,pady=15,padx=10)
 tableb=tk.Button(f2,text="CREATE TABLE",command=create_table,bg="blue",fg="white")
 tableb.pack(side=tk.LEFT,fill=tk.X,pady=15,padx=10)
 f8=tk.Frame(sc)
 f8.pack(fill=tk.BOTH)
 gpfl1=tk.Label(f8,text="Employee_name     :",font=('bold'))
 gpfl1.pack(side=tk.LEFT,fill=tk.X,pady=15,padx=10)
 global gpfno1
 gpfno1=tk.Entry(f8,bg="burlywood",width=51,font=(20))
 gpfno1.pack(side=tk.LEFT,fill=tk.X,pady=15,padx=10)
 f6=tk.Frame(sc)
 f6.pack(fill=tk.BOTH)
 gpfl=tk.Label(f6,text="                                    ",font=('bold'))
 gpfl.pack(side=tk.LEFT,fill=tk.X,pady=15,padx=10)
 op=tk.Button(f2,text="VIEW TABLE",command=opent,bg="blue",fg="white")
 op.pack(side=tk.LEFT,fill=tk.X,pady=15,padx=10)
 op2=tk.Button(f6,text="RESET",command=reset1,bg="blue",fg="white")
 op2.pack(side=tk.LEFT,fill=tk.X,pady=15,padx=10)
 global op3
 op3=tk.Button(f6,text="START",command=start,bg="blue",fg="white")
 op3.pack(side=tk.LEFT,fill=tk.X,pady=15,padx=10)
 op3.config(state=DISABLED)
 readsrno()
 f9=tk.Frame(sc)
 f9.pack(fill=tk.BOTH)
 global count
 count=tk.Label(f9,font=('bold'),text="Count")
 count.pack()
 sc.mainloop()
 return
def readsrno():  
 of=open("srno.ot","r")
 data=of.readline()
 global srno
 srno=data
 of.close()
 book = open_workbook('email.xlsx')
 for s in book.sheets():
  i=0
 global empid,empname
 try:
  empid=s.cell(int(srno),0).value
  empname=s.cell(int(srno),2).value
 except:
  of=open("srno.ot","w")
  of.write("0")
  of.close()
  of=open("srno.ot","r")
  data=of.readline()
  srno=data
  of.close()
  book = open_workbook('email.xlsx')
  for s in book.sheets():
   i=0
  empid=s.cell(int(srno),0).value
  empname=s.cell(int(srno),2).value
  gpfno.delete(0, 'end')
  gpfno.insert(0,empid)
  gpfno1.delete(0, 'end')
  gpfno1.insert(0,empname)
  sc.update()
  return
 gpfno.delete(0, 'end')
 gpfno.insert(0,empid)
 gpfno1.delete(0, 'end')
 gpfno1.insert(0,empname)
 sc.update()
 return
def writesrno():
 global srno
 srno=srno
 srno=int(srno)+1
 of=open("srno.ot","w")
 of.write(str(srno))
 of.close()
 return
def reset1():
 of=open("srno.ot","w")
 of.write("0")
 of.close()
 of=open("srno.ot","r")
 data=of.readline()
 global srno
 srno=data
 of.close()
 book = open_workbook('email.xlsx')
 for s in book.sheets():
  i=0
 global empid
 empid=s.cell(int(srno),0).value
 empname=s.cell(int(srno),2).value
 gpfno.delete(0, 'end')
 gpfno.insert(0,empid)
 gpfno1.delete(0, 'end')
 gpfno1.insert(0,empname)
 count.config(text="0")
 sc.update()
 return
def create_table(fl=0):
 try:
  global f,data
  if fl==1:
    f=file1.get()
  else:
    f=file.get()
  global gpf
  gpf=gpfno.get()
  gpf=gpf.replace(" ","")
  req = urllib.request.Request(f)
  resp = urllib.request.urlopen(req)
  respData = resp.read()
  j=respData.decode("utf-8")
 except:
  mbox.showinfo("Failed","Incorect File!")
  return
 if((j.find(gpf,0,-1)==-1)or(gpf=="")):
  if fl==1:
    mbox.showinfo("Failed","Record Incorect!")
    return
  else:
    create_table(fl=1)
 else:
  i=j.find(gpf,0,-1)
  ios1=0
  while(ios1<i):
   ios=ios1
   t='<table cellpadding="0" cellspacing="0"  align="CENTER" width="100%">'
   ios1=j.find(t,ios1+1,-1)
   if(ios1==-1):
    ios1=i
  k=0
  ioe=j.find('</table>',i,-1)
  while(k<12):
   ioe=j.find('</table>',ioe+1,-1)
   k=k+1
  table=j[ios:ioe]
  fl=0
  y="""<table cellpadding="0" cellspacing="0"  align="CENTER" width="100%">
					<tr>
						<td>&nbsp;</td>
					</tr>
					<tr>
						<td>&nbsp;</td>
					</tr>
					<tr>
						<td style="border-color:#353732;border-bottom: 1px solid;" width="100%">&nbsp;</td>
					</tr>
					<tr>
						<td>&nbsp;</td>
					</tr>
					<tr>
						<td>&nbsp;</td>
					</tr>
				</table>"""
  global html1,html2,coll,dis
  
  cutt=""
  html1="<html><body>"+table+y+"""</table></table></body></html>"""
  html2="""<html><body><table border="1"><tr><th>NON GOVT. CUTTING</th><th>VALUE</th></tr>"""
  try:
   xlsx()
   sc.update()
   global diss
   diss=dis
   if(diss==0):
    return
  except:
   mbox.showinfo("Failed","Chack email.xlsx file!")
   return
  for i in range(coll):
   if(i==0):
     i=0
   else:  
    cutt=cutt+"<tr><td>"+str(col[i])+"</td><td>"+str(data[i])+"</td></tr>"
  html2=html2+cutt+"</table></body></html>"
  file=open("table.html","w")
  file.write(html1)
  file.close()
  file=open("tablec.html","w")
  file.write(html2)
  file.close()
  global msg
  global reset
  reset=1
 return
def opent():
 try:
  os.system("table.html")
  os.system("tablec.html")
 except:
  mbox.showinfo("Failed","Table does not found!")
  return
 return
def logoutt():
 account.entryconfig(0,state=NORMAL)
 account.entryconfig(1,state=DISABLED)
 op3.config(state=DISABLED)
 return
def browse(f):
 if f==1:
   file.config(state='normal')
   file.delete(0, 'end')
 else:
   file1.config(state='normal')
   file1.delete(0, 'end')
 global source_file
 source_file=filedialog.askopenfilename()
 source_file="file:///"+source_file
 if f==1:
   file.insert(0,source_file)
   file.config(state='readonly')
 else:
   file1.insert(0,source_file)
   file1.config(state='readonly')
 return
def browsefb():
 fileb.config(state='normal')
 fileb.delete(0, 'end')
 global source_fileb
 source_fileb=filedialog.askopenfilename()
 fileb.insert(0,source_fileb)
 fileb.config(state='readonly')
 return
def login2():
 global gmail
 gmail=gmail_id.get()
 global passw
 passw=gmail_pass.get()
 try:
  s = smtplib.SMTP('smtp.gmail.com', 587)
  s.starttls()
 except:
  msg.config(text="Failed! Chack Internet Connection!")
  return
 try:
  s.login(gmail,passw)
 except:
  msg.config(text="Failed! Chack Internet Connection! Or Login Details!")
  return
 mbox.showinfo("Success","Login Success!")
 account.entryconfig(1,state=NORMAL)
 account.entryconfig(0,state=DISABLED)
 op3.config(state=NORMAL)
 sc.destroy()
 return
def loginf():
 global sc
 sc=tk.Toplevel()
 sc.title('LOGIN')
 sc.minsize(630,200)
 sc.maxsize(630,200)
 l=tk.Label(sc,text="LOGIN",bg="brown",fg="white",font=('bold'))
 l.pack(fill=tk.X)
 f3=tk.Frame(sc)
 f3.pack(fill=tk.BOTH)
 gl=tk.Label(f3,text="Gmail_ID                  :",font=('bold'))
 gl.pack(side=tk.LEFT,fill=tk.X,pady=15,padx=10)
 global gmail_id
 gmail_id=tk.Entry(f3,bg="burlywood",width=45,font=(20))
 gmail_id.pack(side=tk.LEFT,fill=tk.X,pady=15,padx=10)
 f5=tk.Frame(sc)
 f5.pack(fill=tk.BOTH)
 gpl=tk.Label(f5,text="Gmail_Password    :",font=('bold'))
 gpl.pack(side=tk.LEFT,fill=tk.X,pady=15,padx=10)
 global gmail_pass
 gmail_pass=tk.Entry(f5,bg="burlywood",width=45,font=(20),show="*")
 gmail_pass.pack(side=tk.LEFT,fill=tk.X,pady=15,padx=10)
 gb=tk.Button(sc,text="LOGIN",command=login2,bg="blue",fg="white")
 gb.pack(side=tk.RIGHT,fill=tk.X,pady=15,padx=47)
 f4=tk.Frame(sc)
 f4.pack(fill=tk.BOTH)
 global msg
 msg=tk.Label(f4,text="",font=('bold'),fg="red")
 msg.pack(side=tk.LEFT,fill=tk.X,pady=15,padx=10)
 sc.after(1, lambda:gmail_id.focus_force() )
 sc.bind('<Return>',lambda event:login2())
 sc.mainloop()
 return
def xlsx2():
 global dis,data,coll,col
 dis=1
 book = open_workbook(source_fileb)
 for s in book.sheets():
  i=0
 row=0
 col=s.row_values(2)
 coll=len(col)
 while(row<s.nrows):
  if s.cell(row,1).value ==gpf :
   data=s.row_values(row)
   return
  row=row+1
 mbox.showinfo("Failed","Email record not found for this Employee_ID ! please update Cutting record")
 dis=0
 return  
def xlsx():
 global dis
 dis2=1
 book = open_workbook('email.xlsx')
 for s in book.sheets():
  i=0
 row=0
 while(row<s.nrows):
  if s.cell(row,0).value ==gpf :
   global remail
   remail=s.cell(row,1).value
   xlsx2()
   return
  row=row+1
 mbox.showinfo("Failed","Email record not found for this Employee_ID ! please update Email record")
 dis=0
 return
def start():
 sc.update()
 of=open("srno.ot","r")
 data=of.readline()
 global srno1
 srno1=data
 of.close()
 book = open_workbook('email.xlsx')
 for s in book.sheets():
  i=0
 r=s.nrows
 i=1
 while(int(srno1)<r):
  readsrno()
  sc.update()
  create_table()
  global reset
  if(reset==0):
   return
  try:
   s = smtplib.SMTP('smtp.gmail.com', 587)
   sc.update()
   s.starttls()
   sc.update()
  except:
   mbox.showinfo("Failed","Chack Internet Connection!")
   return
  try:
   s.login(gmail,passw)
   sc.update()
  except:
   mbox.showinfo("Failed","Chack Internet Connection!")
   return
  global remail,month
  monthg=month.get()
  try:
   body=MIMEMultipart('alternative')
   sub='Subject:Payment Details'+monthg+'\n'
   part1=MIMEText(sub,'plain')
   part2=MIMEText(html1,'html')
   body.attach(part1)
   body.attach(part2)
   sc.update()
   remail=remail
   s.sendmail(gmail,remail,'Subject:Payment Details'+monthg+'\n'+body.as_string())
   #s.quit()
   body=MIMEMultipart('alternative')
   sub='Subject:Non Govt. Cutting Details'+monthg+'\n'
   part1=MIMEText(sub,'plain')
   part2=MIMEText(html2,'html')
   body.attach(part1)
   body.attach(part2)
   sc.update()
   s.sendmail(gmail,remail,'Subject:Non Govt. Cutting Details'+monthg+'\n'+body.as_string())
   s.quit()
   sc.update()
   writesrno()
   sc.update()
   i=i+1
   reset=0
   srno1=int(srno1)+1
   count.config(text="Total Sent:"+str(srno1))
   sc.update()
   time.sleep(2)
  except:
   mbox.showinfo("Failed","Email dosn't send!")
   return
 mbox.showinfo("Success","All Sent")
 reset1()
 return
def chpass():
 global sc3
 sc3=tk.Toplevel()
 sc3.title('CHANGE PASSWORD')
 sc3.minsize(630,200)
 sc3.maxsize(630,200)
 l=tk.Label(sc3,text="CHANGE PASSWORD",bg="brown",fg="white",font=('bold'))
 l.pack(fill=tk.X)
 f3=tk.Frame(sc3)
 f3.pack(fill=tk.BOTH)
 gl=tk.Label(f3,text="Old Password      :",font=('bold'))
 gl.pack(side=tk.LEFT,fill=tk.X,pady=15,padx=10)
 global oldpass
 oldpass=tk.Entry(f3,bg="burlywood",width=45,font=(20),show="*")
 oldpass.pack(side=tk.LEFT,fill=tk.X,pady=15,padx=10)
 f5=tk.Frame(sc3)
 f5.pack(fill=tk.BOTH)
 gpl=tk.Label(f5,text="New Password    :",font=('bold'))
 gpl.pack(side=tk.LEFT,fill=tk.X,pady=15,padx=10)
 global newpass
 newpass=tk.Entry(f5,bg="burlywood",width=45,font=(20),show="*")
 newpass.pack(side=tk.LEFT,fill=tk.X,pady=15,padx=10)
 gb=tk.Button(sc3,text="CHANGE",command=change,bg="blue",fg="white")
 gb.pack(side=tk.RIGHT,fill=tk.X,pady=15,padx=47)
 f4=tk.Frame(sc3)
 f4.pack(fill=tk.BOTH)
 global msg2
 msg2=tk.Label(f4,text="",font=('bold'),fg="red")
 msg2.pack(side=tk.LEFT,fill=tk.X,pady=15,padx=10)
 sc3.after(1, lambda:oldpass.focus_force() )
 sc3.bind('<Return>',lambda event:change())
 sc3.mainloop()
 return
def change():
 global oldpass
 passo=oldpass.get()
 global newpass
 passn=newpass.get()
 f=open("run.txt","r")
 p=f.read()
 f.close()
 if passo==p:
   if len(passn)>5:
    f=open("run.txt","w")
    f.write(passn)
    f.close()
    mbox.showinfo("Success","Password Changed!")
    global sc3
    sc3.destroy()
   else:
    mbox.showinfo("Failed","Password Length Less Than 6 !")  
    sc3.focus_set()
 else:
  mbox.showinfo("Failed","Old Password Does Not Match!")
  sc3.focus_set()
 return

main()
#validate()
