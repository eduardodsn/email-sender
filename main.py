import sqlite3
import smtplib
import os.path
import random
from tkinter.ttk import Treeview
from tkinter import *
from tkinter import messagebox
from tkinter.filedialog import askopenfilenames
from PIL import ImageTk, Image
from email import encoders
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart


class EmailSender():
    def __init__(self):
        self.conn = sqlite3.connect('database.db')
        self.cursor = self.conn.cursor()
        self.id_emails_db = []
        self.email_cod = ''
        self.uniq_id = 0
        self.attach = []
        self.to_send_from_db = ''
        self.cod_page = ''
        self.create_db()
        self.search_data_ids()


    def create_db(self):
        self.cursor.execute('CREATE TABLE IF NOT EXISTS data (id INTEGER, email TEXT NOT NULL, name TEXT)')


    def search_data_ids(self):
        self.cursor.execute('SELECT * FROM data')
        rows = self.cursor.fetchall()
        for row in rows:
            self.id_emails_db.append(row[0])


    def nav_menu(self, cod, app):

        app.destroy()
        if cod == 'main':
            try:
                self.to_send_from_db = ''
                self.cod_page = ''
                self.smtp.quit()
            except:
                pass
            self.main_page()
        elif cod == 'single':
            self.single_page()
        elif cod == 'login':
            self.login_page()
        elif cod == 'acess_db':
            self.access_db_page()


    def server_page_save(self, server_entry, app):
        self.server = server_entry.get()
        app.destroy()


    def show_pass(self, pass_entry):
        if pass_entry['show'] == '':
            pass_entry['show'] = '•'
        else:
            pass_entry['show'] = ''


    def centering_page(self, app, app_width, app_height):
        screen_width = app.winfo_screenwidth()
        screen_height = app.winfo_screenheight()

        x_cordinate = int((screen_width/2) - (app_width/2))
        y_cordinate = int((screen_height/2) - (app_height/2))
        app.geometry("{}x{}+{}+{}".format(app_width, app_height, x_cordinate, y_cordinate))


    def attach_files(self, attach_entry):
        files = askopenfilenames()

        for path in files:
            part = MIMEBase('application', "octet-stream")
            with open(path, 'rb') as file:
                part.set_payload(file.read())
            encoders.encode_base64(part)
            file_name = os.path.basename(path)
            attach_entry.insert(END, f'{os.path.basename(file_name)}'+'; ')
            part.add_header('Content-Disposition',
                            'attachment; filename="{}"'.format(file_name))
            self.attach.append(part)


    def login_auth(self, app, email_entry, pass_entry, cod):
        self.login = email_entry.get()
        self.password = pass_entry.get()

        try:
            if cod == 'gmail':
                messagebox.showinfo('Login', 'Connecting...')
                self.smtp = smtplib.SMTP_SSL('smtp.gmail.com', 465)
                self.smtp.login(self.login, self.password)
            elif cod == 'exchange':
                messagebox.showinfo('Login', 'Connecting...')
                self.smtp = smtplib.SMTP(self.server)
                self.smtp.starttls()
                self.smtp.login(self.login, self.password)
        except:
            messagebox.showerror('Login', 'Login failed, try again!')
        else:
            self.nav_menu('single', app)


    def send_email(self, app, to_entry, subject_text, content_text, attach_entry):
        if to_entry.get() == '' or subject_text.get("1.0", END) == '':
            messagebox.showerror('Error', 'No email address to send or empty subject!')
        else:
            to_send = to_entry.get()
            to_send = to_send.split('; ')
            
            for email in to_send:
                if email != '':
                    self.msg = MIMEMultipart()
                    self.msg['From'] = self.login
                    self.msg['To'] = email
                    self.msg['Subject'] = subject_text.get("1.0", END)
                    self.msg.attach(MIMEText(content_text.get("1.0", END)))

                    for part in self.attach:
                        self.msg.attach(part)
                    try:
                        self.smtp.sendmail(self.login, self.msg['To'], self.msg.as_string())
                    except:
                        messagebox.showerror('Error', 'Cannot send the email!')
                    else:
                        if self.cod_page != 'from_db':
                            messagebox.showinfo('Info', 'Email sent to {} successfully!'.format(email))

            self.attach.clear()

            to_entry.delete(0, END)
            attach_entry.delete(0, END)
            subject_text.delete(1.0, END)
            content_text.delete(1.0, END)                        
            
            if self.cod_page == 'from_db':
                messagebox.showinfo('Info', 'Emails sent successfully!')
            app.destroy()
            self.restart_nav_page()


    def restart_nav_page(self):
        app = Tk()
        self.centering_page(app, 500, 300)
        app.title('Nav Page')
        app.resizable(False, False)
        app['background'] = '#075fab'
        app.iconbitmap('icons/001-restart.ico')

        nav_label = Label(app, text='Navigation Page', font=('Terminal', 23), bg='#075fab', fg='White')
        nav_label.place(relx=0.5, rely=0.15, anchor=CENTER)

        back_main_btn = Button(app, text='Back to Main Page', font=('Roboto', 13), bg='Gray25',
                        fg='White', command=lambda: self.nav_menu('main', app))
        back_main_btn.place(relx=0.5, rely=0.4, anchor=CENTER, width=160, height=50)

        send_again_btn = Button(app, text='Send email again', font=('Roboto', 13), bg='Green',
                        fg='White', command=lambda: self.nav_menu('single', app))
        send_again_btn.place(relx=0.5, rely=0.6, anchor=CENTER, width=160, height=50)

        exit_btn = Button(app, text='Exit', font=('Roboto', 13), bg='Red', fg='White', 
                                    command=lambda: self.nav_menu('', app))
        exit_btn.place(relx=0.5, rely=0.8, anchor=CENTER, width=160, height=50)
        app.mainloop()


    def server_page(self):
        app = Toplevel()
        self.centering_page(app, 400, 150)
        app.title('Server')
        app.iconbitmap('icons/002-server.ico')
        app['background'] = '#075fab'
        app.resizable(False, False)

        server_label = Label(app, text='Exchange server:', font=('Roboto', 12), bg='#075fab', fg='White')
        server_label.place(relx=0.18, rely=0.4, anchor=CENTER)

        server_entry = Entry(app, font=('Roboto', 12), width=27)
        server_entry.place(relx=0.65, rely=0.4, anchor=CENTER)

        save_btn = Button(app, text='Save', font=('Roboto', 12), width=8, bg='Green', fg='White',
                                    command=lambda: self.server_page_save(server_entry, app))
        save_btn.place(relx=0.5, rely=0.7, anchor=CENTER)

        app.mainloop()


    def search_all_data_db(self, tree):
        tree.delete(*tree.get_children())
        self.cursor.execute('SELECT * FROM data')
        rows = self.cursor.fetchall()
        for row in rows:
            tree.insert("" , "end", values=(f'{row[0]}', f'{row[1]}',f'{row[2]}'))


    def fill_tree(self, tree):
        rows = self.cursor.fetchall()
        if not rows:
            tree.delete(*tree.get_children())
            messagebox.showerror('Error', "Didn't find data!")
        else:
            tree.delete(*tree.get_children())
            for row in rows:
                tree.insert("" , "end", values=(f'{row[0]}', f'{row[1]}',f'{row[2]}'))  


    def search_one_data_db(self, name_entry, email_entry, tree):
        if name_entry.get() == '' and email_entry.get() == '':
            tree.delete(*tree.get_children())
            self.search_all_data_db(tree)
            return
        else:
            if name_entry.get() != '': # name field not empty
                    self.cursor.execute('SELECT * FROM data WHERE name=?', [name_entry.get()])
                    self.fill_tree(tree)
                      
            else: # name field empty, so use email_entry
                self.cursor.execute('SELECT * FROM data WHERE email=?', [email_entry.get()])
                self.fill_tree(tree)        

    
    def search_data_db(self, cod, name_entry, email_entry, tree):
        if cod == 'all':
            self.search_all_data_db(tree)
        elif cod == 'one':
            self.search_one_data_db(name_entry, email_entry, tree)                
    

    def delete_data_db(self, email_entry, tree):
        if not tree.focus():
            messagebox.showinfo('Info', 'Not a data selected!')
        else:
            item = tree.focus()
            try:
                id_item = tree.item(item)['values'][0]
                email_item = tree.item(item)['values'][1]
            except:
                pass

            if messagebox.askquestion('Delete data', 'Are you sure you want to delete this data?') == 'yes':
                self.cursor.execute('DELETE FROM data WHERE email=?', [email_item])
                self.conn.commit()
                self.id_emails_db.remove(id_item)
                self.search_all_data_db(tree)

    
    def insert_data_db(self, name_entry, email_entry, tree):
        email = email_entry.get()
        name = name_entry.get()
        
        if name_entry.get() == '' or email_entry.get() == '':
            messagebox.showerror('Error', 'Email or Name field empty!')
            return

        if self.email_cod == '' and self.uniq_id == 0:
            self.uniq_id = random.randint(1, 1000001)
            
            while True:
                if self.uniq_id in self.id_emails_db:
                    self.uniq_id = random.randint(1, 1000001)
                else:
                    self.id_emails_db.append(self.uniq_id)
                    break

            try:
                self.cursor.execute("INSERT INTO data (id, email, name) VALUES (?, ?, ?)", (self.uniq_id, email, name))
                self.conn.commit()
            except:
                messagebox.showerror('Error', 'Impossible to insert on DB!')
                return
        elif self.email_cod == 'edit':
            try:
                self.cursor.execute("UPDATE data SET email=?, name=? WHERE id=?", (email, name, self.uniq_id))
                self.conn.commit()
            except:
                messagebox.showerror('Error', 'Impossible to update on DB!')
                return
        
        messagebox.showinfo('Info', 'Data saved successfully!')
        name_entry.delete(0, END)
        email_entry.delete(0, END)
        self.search_all_data_db(tree)
        self.email_cod = ''
        self.uniq_id = 0
    

    def edit_data_db(self, name_entry, email_entry, tree):
        if not tree.focus():
            messagebox.showinfo('Info', 'Not a data selected!')
        else:
            self.email_cod = 'edit'
            item = tree.focus()
            self.uniq_id= tree.item(item)['values'][0]
            email_item = tree.item(item)['values'][1]
            name_item = tree.item(item)['values'][2]
            email_entry.insert(END, email_item)
            name_entry.insert(END, name_item)


    def append_one_db(self, added_email_entry, tree):
        if not tree.focus():
            messagebox.showinfo('Info', 'Not a data selected!')
        else:
            added_email_entry.delete(0, END)
            item = tree.focus()
            email_item = tree.item(item)['values'][1]
            added_email_entry.insert(END, email_item + '; ')


    def append_all_db(self, added_email_entry, tree):
        added_email_entry.delete(0, END)
        self.cursor.execute('SELECT * FROM data')
        rows = self.cursor.fetchall()
        for row in rows:
            added_email_entry.insert(END, row[1] + '; ')


    def clear_attach_files(self, attach_entry):
        self.attach.clear()
        attach_entry.delete(0, END)


    def send_email_from_db(self, added_email_entry, app):
        self.cod_page = 'from_db'
        if added_email_entry.get() == '':
            messagebox.showerror('Error', 'No email address to send!')
            return
        self.to_send_from_db = added_email_entry.get()
        app.destroy()
        if self.login_page():
            self.single_page()


    def access_db_page(self):
        app = Tk()
        self.centering_page(app, 700, 550)
        app.title('Database')
        app.iconbitmap('icons/003-database.ico')
        app['background'] = '#075fab'
        app.resizable(False, False)
        
        db_label = Label(app, text='Database', font=('Terminal', 25), bg='#075fab', fg='White')
        db_label.place(relx=0.5, rely=0.1, anchor=CENTER)

        email_label = Label(app, text='Email:', font=('Roboto', 12), bg='#075fab', fg='White')
        email_label.place(relx=0.1, rely=0.25, anchor=CENTER)

        email_entry = Entry(app, font=('Roboto', 13), width=45)
        email_entry.place(relx=0.14, rely=0.25, anchor='w')

        name_label = Label(app, text='Name:', font=('Roboto', 12), bg='#075fab', fg='White')
        name_label.place(relx=0.1, rely=0.32, anchor=CENTER)

        name_entry = Entry(app, font=('Roboto', 13), width=45)
        name_entry.place(relx=0.14, rely=0.32, anchor='w')

        search_image = Image.open('icons/004-search.ico').resize((20, 20), Image.ANTIALIAS)
        search_image = ImageTk.PhotoImage(search_image)
        search_btn = Button(app, image=search_image, compound=CENTER, width=27, height=22, bg='White',
                                    command=lambda: self.search_data_db('one', name_entry, email_entry, tree))
        search_btn.place(relx=0.77, rely=0.25, anchor='w')
        
        edit_image = Image.open('icons/014-edit.ico').resize((30, 30), Image.ANTIALIAS)
        edit_image = ImageTk.PhotoImage(edit_image)
        edit_btn = Button(app, image=edit_image, compound=CENTER, width=27, height=22, bg='White',
                                    command=lambda: self.edit_data_db(name_entry, email_entry, tree))
        edit_btn.place(relx=0.86, rely=0.25, anchor='w')

        delete_image = Image.open('icons/005-delete.ico').resize((22, 22), Image.ANTIALIAS)
        delete_image = ImageTk.PhotoImage(delete_image)
        delete_btn = Button(app, image=delete_image, compound=CENTER,  width=27, height=22, bg='White',
                                    command=lambda: self.delete_data_db(email_entry, tree))
        delete_btn.place(relx=0.86, rely=0.32, anchor='w')

        save_image = Image.open('icons/012-save.ico').resize((20, 20), Image.ANTIALIAS)
        save_image = ImageTk.PhotoImage(save_image)
        save_btn = Button(app, image=save_image, width=27, height=22, bg='White', 
                                    command=lambda: self.insert_data_db(name_entry, email_entry, tree))
        save_btn.place(relx=0.77, rely=0.32, anchor='w')

        append_one_btn = Button(app, text='Append', font=('Roboto', 10), width=8, height=1, bg='White', fg='Black',
                                    command=lambda: self.append_one_db(added_email_entry, tree))
        append_one_btn.place(relx=0.8, rely=0.82, anchor='w')

        append_all_btn = Button(app, text='Append all', font=('Roboto', 10), width=8, height=1, bg='White', fg='Black',
                                    command=lambda: self.append_all_db(added_email_entry, tree))
        append_all_btn.place(relx=0.8, rely=0.88, anchor='w')

        added_email_label = Label(app, text='Added:', font=('Roboto', 12), bg='#075fab', fg='White')
        added_email_label.place(relx=0.1, rely=0.95, anchor=CENTER)

        added_email_entry = Entry(app, font=('Roboto', 12), width=50)
        added_email_entry.place(relx=0.14, rely=0.95, anchor='w')

        send_email_btn = Button(app, text='Send email', font=('Roboto', 10), width=8, height=1, bg='Green', fg='White',
                                    command=lambda: self.send_email_from_db(added_email_entry, app))
        send_email_btn.place(relx=0.8, rely=0.95, anchor='w')

        back_icon = Image.open('icons/006-return.ico').resize((25, 25), Image.ANTIALIAS)
        back_icon = ImageTk.PhotoImage(back_icon)
        back_btn = Button(app, image=back_icon, compound=CENTER, width=25, height=25, bg='White',
                                    command=lambda: self.nav_menu('main', app))
        back_btn.place(relx=0.05, rely=0.1, anchor=CENTER)

        frame = Frame(app, bg='White')
        frame.pack()
        frame.place(relx=0.49, rely=0.57, anchor=CENTER)

        tree = Treeview(frame, height=10)
        tree.pack(side=TOP, fill=NONE, expand=FALSE)

        tree['columns'] = ('index', 'email', 'name')
        tree.column('#0', width=0, stretch=NO)
        tree.column('index', width=80, stretch=NO)
        tree.column('email', width=250, stretch=NO)
        tree.column('name', width=255, stretch=NO)

        tree.heading('index', text='Index', anchor=CENTER)
        tree.heading('email', text='Email', anchor=CENTER)
        tree.heading('name', text='Name', anchor=CENTER)
        self.search_data_db('all', name_entry, email_entry, tree)

        app.mainloop()


    def login_page(self):
        app = Tk()
        self.centering_page(app, 500, 250)
        app.iconbitmap('icons/007-login.ico')
        app.title('Login')
        app['background'] = '#075fab'
        app.resizable(False, False)

        info_label = Label(app, text='Log in to send emails', font=('Terminal', 12), bg='#075fab', fg='White')
        info_label.place(relx=0.5, rely=0.1, anchor=CENTER)

        email_label = Label(app, text='Email:', font=('Roboto', 12), bg='#075fab', fg='White')
        email_label.place(relx=0.15, rely=0.3, anchor='w')

        email_entry = Entry(app, font=('Roboto', 12), fg='black', width=32)
        email_entry.place(relx=0.25, rely=0.3, anchor='w')

        pass_label = Label(app, text='Password:', font=('Roboto', 12),bg='#075fab', fg='White')
        pass_label.place(relx=0.09, rely=0.45, anchor='w')

        pass_entry = Entry(app, font=('Roboto', 12), show='•', fg='black', width=32)
        pass_entry.place(relx=0.25, rely=0.45, anchor='w')
        
        show_pass_icon = Image.open('icons/013-show-pass.ico').resize((25, 25), Image.ANTIALIAS)
        show_pass_icon = ImageTk.PhotoImage(show_pass_icon)

        show_pass_btn = Button(app, image=show_pass_icon, compound=CENTER, height=18, width=22, bg='White',
                                    command=lambda: self.show_pass(pass_entry))
        show_pass_btn.place(relx=0.85, rely=0.45, anchor='w')
    
        selected_radio = StringVar()

        gmail_radio = Radiobutton(app, text='Gmail', font=('Roboto', 11), value='gmail', fg='White', bg='#075fab', activebackground='#075fab',
                    activeforeground='White', selectcolor='#075fab', variable=selected_radio)
        gmail_radio.place(relx=0.3, rely=0.58, anchor='w')
        gmail_radio.select()

        exchange_radio = Radiobutton(app, text='Exchange', font=('Roboto', 11), value='exchange', fg='White', bg='#075fab', activebackground='#075fab',
                    activeforeground='White', selectcolor='#075fab', variable=selected_radio,command=lambda: self.server_page())
        exchange_radio.place(relx=0.55, rely=0.58, anchor='w')

        login_btn = Button(app, text='Login', font=('Roboto', 12), height=1, width=6, bg='White', fg='Black',
                                    command=lambda: self.login_auth(app, email_entry, pass_entry, selected_radio.get()))
        login_btn.place(relx=0.5, rely=0.75, anchor=CENTER)

        back_icon = Image.open('icons/006-return.ico').resize((25, 25), Image.ANTIALIAS)
        back_icon = ImageTk.PhotoImage(back_icon)
        back_btn = Button(app, image=back_icon, compound=CENTER, width=25, height=25, bg='White',
                                    command=lambda: self.nav_menu('main', app))
        back_btn.place(relx=0.05, rely=0.1, anchor=CENTER)

        app.mainloop()


    def single_page(self):
        app = Tk()
        self.centering_page(app, 700, 450)
        app.iconbitmap('icons/008-send.ico')
        app.title('Send a single email')
        app['background'] = '#075fab'
        app.resizable(False, False)

        single_title_label = Label(app, text='Send a single email', font=('Terminal', 25), bg='#075fab', fg='White')
        single_title_label.place(relx=0.5, rely=0.1, anchor=CENTER)

        to_label = Label(app, text='To:', font=('Arial', 14), bg='#075fab', fg='White')
        to_label.place(relx=0.15, rely=0.25, anchor='e')

        to_entry = Entry(app, font=('Arial', 11), fg='black', width=50)
        to_entry.place(relx=0.16, rely=0.25, anchor='w')

        if self.cod_page == 'from_db':
            to_entry.delete(0, END)
            to_entry.insert(END, self.to_send_from_db)

        subject_label = Label(app, text='Subject:', font=('Arial', 14), bg='#075fab', fg='White')
        subject_label.place(relx=0.15, rely=0.35, anchor='e')

        subject_text = Text(app, font=('Arial', 11), height=1, width=50)
        subject_text.place(relx=0.16, rely=0.33, anchor='nw')

        content_label = Label(app, text='Content:', font=('Arial', 14), bg='#075fab', fg='White')
        content_label.place(relx=0.15, rely=0.45, anchor='e')

        content_text = Text(app, font=('Arial', 11), height=7, width=50)
        content_text.place(relx=0.16, rely=0.43, anchor='nw')

        attach_label = Label(app, text='Attach files:', font=('Arial', 14), bg='#075fab', fg='White')
        attach_label.place(relx=0.15, rely=0.75, anchor='e')

        attach_entry = Entry(app, font=('Arial', 11), width=50)
        attach_entry.place(relx=0.16, rely=0.73, anchor='nw')

        select_attach_btn = Button(app, text='Select', font=('Arial', 11), height=1, width=6, bg='White', fg='Black',
                                    command=lambda: self.attach_files(attach_entry))
        select_attach_btn.place(relx=0.75, rely=0.72, anchor='nw')

        clear_attach_btn = Button(app, text='Clear Attach', font=('Arial', 11), height=1, width=9, bg='White', fg='Black',
                                    command=lambda: self.clear_attach_files(attach_entry))
        clear_attach_btn.place(relx=0.85, rely=0.72, anchor='nw')
        
        send_btn = Button(app, text='Send email', font=('Roboto', 11), height=1, width=8, bg='green', fg='White',
                                    command=lambda: self.send_email(app, to_entry, subject_text, content_text, attach_entry))
        send_btn.place(relx=0.5, rely=0.86, anchor=CENTER)

        back_icon = Image.open('icons/006-return.ico').resize((25, 25), Image.ANTIALIAS)
        back_icon = ImageTk.PhotoImage(back_icon)
        back_btn = Button(app, image=back_icon, compound=CENTER, width=25, height=25, bg='White',
                                    command=lambda: self.nav_menu('main', app))
        back_btn.place(relx=0.05, rely=0.1, anchor=CENTER)

        app.mainloop()


    def main_page(self):
        app = Tk()
        self.centering_page(app, 700, 450)
        app.iconbitmap('icons/009-home.ico')
        app.title('Email Sender')
        app['background'] = '#075fab'
        app.resizable(False, False)

        app_title_label = Label(app, text='Email', font=('Terminal', 50), bg='#075fab', fg='White')
        app_title_label.place(relx=0.4, rely=0.2, anchor=CENTER)

        app_title_label2 = Label(app, text='Sender', font=('Terminal', 50), bg='#075fab', fg='White')
        app_title_label2.place(relx=0.6, rely=0.37, anchor=CENTER)

        app_image = Image.open('icons/010-mail.ico').resize((100, 100), Image.ANTIALIAS)
        app_image = ImageTk.PhotoImage(app_image)
        app_image_label = Label(app, image=app_image, bg='#075fab')
        app_image_label.place(relx=0.65, rely=0.18, anchor=CENTER)
        
        single_icon = Image.open('icons/008-send.ico').resize((40, 40), Image.ANTIALIAS)
        single_icon = ImageTk.PhotoImage(single_icon)
        single_btn = Button(app, text='  Send a single\n email', image=single_icon, compound=LEFT, font=('Roboto', 14), bg='White', fg='Black', height=75, width=180,
                                    command=lambda: self.nav_menu('login', app))
        single_btn.place(relx=0.2, rely=0.7, anchor=CENTER)

        access_db_icon = Image.open('icons/003-database.ico').resize((30, 30), Image.ANTIALIAS)
        access_db_icon = ImageTk.PhotoImage(access_db_icon)
        access_db_btn = Button(app, text='  Access emails\n DB', image=access_db_icon, compound=LEFT, font=('Roboto', 14),
                            bg='White', fg='Black', height=75, width=180, command=lambda: self.nav_menu('acess_db', app))
        access_db_btn.place(relx=0.5, rely=0.7, anchor=CENTER)

        exit_icon = Image.open('icons/011-exit.ico').resize((30, 30), Image.ANTIALIAS)
        exit_icon = ImageTk.PhotoImage(exit_icon)
        exit_btn = Button(app, text='  Exit', image=exit_icon, compound=LEFT, font=('Roboto', 14), bg='White', fg='Black', height=75, width=180,
                                    command=lambda: self.nav_menu('', app))
        exit_btn.place(relx=0.8, rely=0.7, anchor=CENTER)
        
        app.mainloop()


if __name__ == "__main__":
    email = EmailSender()
    email.main_page()
