from tkinter import *
from PIL import ImageTk
from tkinter import messagebox,filedialog
import os
import pandas as pd
import emailsfunc
import time

class Bulk_Email:
    def __init__(self, root):
        self.root = root
        self.root.title("Bulk Email Application")
        self.root.geometry("1000x600+200+50")
        self.root.resizable(False,False)
        self.root.config(bg="white")

        self.email_icon=ImageTk.PhotoImage(file="images/email.png")
        self.setting_icon=ImageTk.PhotoImage(file="images/setting.png")

        title=Label(self.root, text="Bulk Email Send Panel", image=self.email_icon, padx=10, compound=LEFT,font=("Goudy Old Style",48,"bold"), bg="#222A35", fg="white", anchor="w").place(x=0,y=0, relwidth=1)

        desc=Label(self.root, text="Use Excel File to Send   Bulk email. Ensure the email column name must be email",font=("Calibri (Body)",14), bg="#FFD966", fg="#262626",anchor="w").place(x=0,y=80, relwidth=1)

        btn_setting=Button(self.root, image=self.setting_icon, bd=0, activebackground="#222A35", bg="#222A35",cursor="hand2",command=self.setting_window).place(x=900, y=5)

        self.var_choice=StringVar()

        single=Radiobutton(self.root,text="Single",value="single", variable=self.var_choice,activebackground="white",font=("Times new roman", 30, "bold"),bg="white",fg="#262626",command=self.check_single_or_bulk).place(x=50,y=150)
        bulk=Radiobutton(self.root, text="Bulk",value="bulk",variable=self.var_choice,activebackground="white", font=("Times new roman", 30, "bold"),bg="white",fg="#262626",command=self.check_single_or_bulk).place(x=250,y=150)
        self.var_choice.set("single")

        to=Label(self.root, text="To (Email Id)", font=("times new roman", 18),bg="white").place(x=50,y=250)
        subj=Label(self.root, text="SUBJECT", font=("times new roman", 18),bg="white").place(x=50,y=300)
        msg=Label(self.root, text="MESSAGE", font=("times new roman", 18),bg="white").place(x=50,y=350)

        self.txt_to = Entry(self.root, font=("times new roman", 14),bg="lightyellow")
        self.txt_to.place(x=300,y=250,width=350,height=30)

        self.btn_browse = Button(self.root,command=self.browse_file, text="Browse", font=("times new roman", 14, "bold"), bg="#8FAADC",fg="#262626",activebackground="#8FAADC",activeforeground="#262626", cursor="hand2",state=DISABLED)
        self.btn_browse.place(x=670,y=250,width=120,height=30)

        self.txt_subj = Entry(self.root, font=("times new roman", 14),bg="lightyellow")
        self.txt_subj.place(x=300,y=300,width=450,height=30)

        self.txt_msg = Text(self.root, font=("times new roman", 12),bg="lightyellow")
        self.txt_msg.place(x=300,y=350,width=650,height=120)

        #----------status-----------
        self.lbl_total=Label(self.root, font=("times new roman", 18),bg="white")
        self.lbl_total.place(x=50,y=490)

        self.lbl_sent=Label(self.root, font=("times new roman", 18),bg="white",fg='green')
        self.lbl_sent.place(x=300,y=490)

        self.lbl_left=Label(self.root, font=("times new roman", 18),bg="white",fg='orange')
        self.lbl_left.place(x=420,y=490)

        self.lbl_failed=Label(self.root, font=("times new roman", 18),bg="white",fg='red')
        self.lbl_failed.place(x=550,y=490)

        btn_clear = Button(self.root, text="CLEAR", command=self.clear1, font=("times new roman", 14, "bold"), bg="#262626",fg="white",activebackground="#262626",activeforeground="white",cursor="hand2").place(x=700,y=500,width=120,height=30)
        btn_send = Button(self.root,command=self.send_email, text="SEND", font=("times new roman", 14, "bold"), bg="#00B0F0",fg="white",activebackground="#00B0F0",activeforeground="white", cursor="hand2").place(x=830,y=500,width=120,height=30)
        self.check_file_exist()


    def browse_file(self):
        op=filedialog.askopenfile(initialdir='/',title='Select Excel File For Emails',filetypes=(("All Files","*.*"),("Excel Files",".xlsx")))
        if op != None:
            data=pd.read_excel(op.name)
            if 'Email' in data.columns:
                self.emails=list(data['Email'])
                c=[]
                for i in self.emails:
                 if pd.isnull(i)==False:
                    c.append(i)
                self.emails=c
                if len(self.emails)>0:
                    self.txt_to.config(state=NORMAL)
                    self.txt_to.delete(0,END)
                    self.txt_to.insert(0, str(op.name.split("/")[-1]))
                    self.txt_to.config(state='readonly')
                    self.lbl_total.config(text="Total: "+str(len(self.emails)))
                    self.lbl_sent.config(text="Sent: ")
                    self.lbl_left.config(text="Left: ")
                    self.lbl_failed.config(text="Failed: ")
                else:
                    messagebox.showerror("Error","This does not have any emails", parent=self.root)

            else:
                messagebox.showerror("Error","Please select the file", parent=self.root)


    def send_email(self):
        x=len(self.txt_msg.get('1.0',END))
        
        if self.txt_to.get() == "" or self.txt_subj.get() == "" or x==1:
            messagebox.showerror("Error","All fileds are required",parent=self.root)
        else:
            if self.var_choice.get()=='single':
                status=emailsfunc.email_send_func(self.txt_to.get(),self.txt_subj.get(),self.txt_msg.get('1.0',END),self.from_,self.pass_)
                if status=="s":
                    messagebox.showinfo("Success","Email Sent",parent=self.root)
                if status=="f":
                    messagebox.showerror("Failed","Email Not Sent. Try again",parent=self.root)    
            if self.var_choice.get()=='bulk':
                self.failed=[]
                self.s_count=0
                self.f_count=0
                for x in self.emails:
                    status=emailsfunc.email_send_func(x,self.txt_subj.get(),self.txt_msg.get('1.0',END),self.from_,self.pass_)
                    if status=='s':
                        self.s_count+=1
                    if status=='f':
                        self.f_count+=1   

                    self.status_bar()    
                    time.sleep(1) 

                messagebox.showinfo("Success","Email Sent. Please check Status",parent=self.root)        


    def status_bar(self):
        self.lbl_total.config(text="Status: "+str(len(self.emails))+"=>>")
        self.lbl_sent.config(text="Sent: "+str(self.s_count))
        self.lbl_left.config(text="Left: "+str(len(self.emails)-(self.s_count+self.f_count)))
        self.lbl_failed.config(text="Failed: "+str(self.f_count))
        self.lbl_total.update()
        self.lbl_sent.update()
        self.lbl_left.update()
        self.lbl_failed.update()


    def check_single_or_bulk(self):
        if self.var_choice.get() == "single":
            self.btn_browse.config(state=DISABLED)
            self.txt_to.config(state=NORMAL)
            self.txt_to.delete(0,END)
            self.clear1()

        if self.var_choice.get() == "bulk":
            self.btn_browse.config(state=NORMAL)   
            self.txt_to.delete(0,END) 
            self.txt_to.config(state='readonly')
            

    def clear1(self):
        self.txt_to.config(state=NORMAL)
        self.txt_to.delete(0,END)
        self.txt_subj.delete(0,END)
        self.txt_msg.delete('1.0',END)
        self.var_choice.set('single')
        self.btn_browse.config(state=DISABLED)
        self.lbl_total.config(text="")
        self.lbl_sent.config(text="")
        self.lbl_left.config(text="")
        self.lbl_failed.config(text="")


    def setting_window(self):
        self.check_file_exist()
        self.root2=Toplevel()
        self.root2.title("Setting")
        self.root2.geometry("700x350+350+90")
        self.root2.focus_force()
        self.root2.grab_set()
        self.root2.config(bg="white")

        title2=Label(self.root2, text="Credentials Setting", image=self.setting_icon, padx=10, compound=LEFT,font=("Goudy Old Style",48,"bold"), bg="#222A35", fg="white", anchor="w").place(x=0,y=0, relwidth=1)

        desc2=Label(self.root2, text="Enter email address and password from which you want to send mails",font=("Calibri (Body)",14), bg="#FFD966", fg="#262626",anchor="w").place(x=0,y=80, relwidth=1)

        from_=Label(self.root2, text="Email Id", font=("times new roman", 18),bg="white").place(x=50,y=150)
        pass_=Label(self.root2, text="Password", font=("times new roman", 18),bg="white").place(x=50,y=200)

        self.txt_from = Entry(self.root2, font=("times new roman", 14),bg="lightyellow")
        self.txt_from.place(x=250,y=150,width=330,height=30)

        self.txt_pass = Entry(self.root2, font=("times new roman", 14),bg="lightyellow",show='*')
        self.txt_pass.place(x=250,y=200,width=330,height=30)

        btn_clear2 = Button(self.root2,command=self.clear2, text="CLEAR", font=("times new roman", 14, "bold"), bg="#262626",fg="white",activebackground="#262626",activeforeground="white",cursor="hand2").place(x=300,y=260,width=120,height=30)
        btn_save = Button(self.root2,command=self.save_setting, text="SAVE", font=("times new roman", 14, "bold"), bg="#00B0F0",fg="white",activebackground="#00B0F0",activeforeground="white", cursor="hand2").place(x=430,y=260,width=120,height=30)
        self.txt_from.insert(0,self.from_)
        self.txt_pass.insert(0,self.pass_)


    def clear2(self):
        self.txt_from.delete(0,END)
        self.txt_pass.delete(0,END)

    def check_file_exist(self):
        if os.path.exists('important.txt')==False:
            f=open('important.txt','w')
            f.write(',')
            f.close()    
        f2=open('important.txt','r')
        self.credentials=[]
        for i in f2:
            # print(i)    
            self.credentials.append([i.split(',')[0],i.split(',')[1]])
        # print(self.credentials)   
        self.from_=self.credentials[0][0]  
        self.pass_=self.credentials[0][1]
        # print(self.from_,self.pass_)  


    def save_setting(self):
        if self.txt_from.get()=="" or self.txt_pass.get()=="":
            messagebox.showerror("Error","All fileds are required",parent=self.root2)
        else:
            f=open('important.txt','w')
            f.write(self.txt_from.get()+','+self.txt_pass.get())
            f.close()    
            messagebox.showinfo("Successfully","Saved",parent=self.root2)   
            self.check_file_exist()
            

root=Tk()
obj = Bulk_Email(root)
root.mainloop()