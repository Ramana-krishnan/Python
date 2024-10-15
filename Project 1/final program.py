import tkinter as tk
from tkinter import Toplevel, Label, Entry, Button, OptionMenu, StringVar, Text, END,filedialog
import sqlite3
import pandas as pd
from datetime import datetime

filename = filedialog.askopenfilename(filetypes=[("Excel files", ".xlsx")])

# Extracting the file and converting it into the 
conn = sqlite3.connect("database_file.db")
df = pd.read_excel(filename)
df.to_sql("table1", conn, if_exists="replace", index=False)
curs = conn.cursor()

def application():
    #usernames and passwords
    pass_file=pd.read_excel("password.xlsx")
    op_pass={row['Username']:row['Password'] for i,row in pass_file.iterrows()}
    
    user_pass = {
        "admin":"admin@123",
        "guest":"",
        "Guest":"",
        "Admin":"admin@123"
    }
    user_pass.update(op_pass)
    operators=list(pass_file["Username"])
    
    
    #login
    def login():
        
        username = username_entry.get()
        password = password_entry.get()
        
        if username in user_pass and str(user_pass[username]) == password:
            s.destroy()
            main_application(username)
        else:
            error_label.config(text="Invalid username or password")
    
    # Function to create the main application window
    
    
    def main_application(username):
        r = tk.Tk()
    
       # selected = StringVar()
        cont = StringVar()
        sub_cont = StringVar()
        sub_cont.set("Select")
    
        # Initialize a global variable
        global box
        box = None
    
        def on_select(*args):
            global box
            selected_option = cont.get()
            if selected_option not in ["Product Name", "Product ID", "Employee Name", "Employee ID"]:
                if box:
                    box.destroy()  # Destroy if exist
                box = Entry(r)
                box.place(x=5, y=94)
            else:
                if box:
                    box.destroy()  # Destroy if exists
                sub_options = df[selected_option].unique()
                box = OptionMenu(r, sub_cont, *sub_options)
                box.place(x=5, y=92)
    
        def retrieve():
            T.delete("1.0", END)
            column = cont.get()
            val = sub_cont.get() if column in [
                "Product Name", "Product ID", "Employee Name", "Employee ID"] else box.get()
            display_query = f"SELECT * FROM table1 WHERE \"{column}\"='{val}'"
        
            # Fetch the data
            data = curs.execute(display_query).fetchall()
            
            # Fetch the column names
            column_names = [description[0] for description in curs.description]
            
            # Format and insert column names
            formatted_columns = "{:<15}" * len(column_names)
            T.insert("end", formatted_columns.format(*column_names) + "\n")
            T.insert("end", "-" * 15 * len(column_names) + "\n")
        
            if data:
                for row in data:
                    formatted_row = "{:<15}" * len(row)
                    T.insert("end", formatted_row.format(*row) + "\n")
            else:
                T.insert("end", "No Records Found\n")

    
        def add_data():
            addwindow = Toplevel()
            addwindow.title("Data Entry")
            entries = {}
            option_menus = {}
    
            def autofill_name(*args):
                selected_id = prod_id_var.get()
                product_name = df[df["Product ID"] ==selected_id]["Product Name"].values[0]
                prod_name_var.set(product_name)
    
            def autofill_emp_name(*args):
                selected_emp_id = emp_id_var.get()
                employee_name = df[df["Employee ID"] ==selected_emp_id]["Employee Name"].values[0]
                emp_name_var.set(employee_name)
    
            def insert_data():
                values = {label: entry.get() if label not in option_menus else entry.get()
                          for label, entry in entries.items()}
                values["Uploaded by"] = username
               
                values["Uploaded Time"] = datetime.now().strftime(
                    "%Y-%m-%d %H:%M")  # Adding current date and time
                #username as "Uploaded by"
                a = []
                for i, j in values.items():
                    a.append(j)
                update_query = "INSERT INTO table1 VALUES{0}".format(tuple(a))
                if a[0]!='' and a[1]!="Select" and a[7]!="Select":
                
                    curs.execute(update_query)
                    conn.commit()
                
                    sql = pd.read_sql("SELECT * FROM table1", conn)
                    sql.to_excel(filename, index=False)
                    addwindow.destroy()
                    T.delete("1.0", "end")
                    T.insert("1.0", "SAVED SUCCESSFULLY!!!!!!!!!!!!!!!!!")
                else:
                    pass
                
                            
    
            labels = [
                "S no", "Product ID", "Product Name", "InDC", "In Date",
                "OutDC", "Out Date", "Employee ID","Employee Name", 
                "Defected Component", "Problem", "Reason", "Uploaded by", "Uploaded Time"
            ]
    
            prod_id_var = StringVar()
            prod_id_var.set("Select")
            prod_name_var = StringVar()
            emp_id_var = StringVar()
            emp_id_var.set("Select")
            emp_name_var = StringVar()
    
            dropdown_labels = ["Product ID", "Employee ID"]
    
            for idx, label in enumerate(labels):
                if label not in ["Uploaded by", "Uploaded Time"]:
                    row, col = divmod(idx, 1)
                    Label(addwindow, text=label).grid(row=row+1, column=col*2)
                    if label in dropdown_labels:
                        unique_values = df[label].unique()
                        var = prod_id_var if label == "Product ID" else emp_id_var
                        option_menu = OptionMenu(addwindow, var, *unique_values)
                        option_menu.grid(row=row+1, column=col*2+1)
                        entries[label] = var
                        option_menus[label] = option_menu
                        if label == "Product ID":
                            var.trace('w', autofill_name)
                        elif label == "Employee ID":
                            var.trace('w', autofill_emp_name)
                    else:
                        entry = Entry(addwindow)
                        entry.grid(row=row+1, column=col*2+1)
                        entries[label] = entry
                        if label == "Product Name":
                            entries[label] = Entry(addwindow, textvariable=prod_name_var)
                            entries[label].grid(row=row+1, column=col*2+1)
                        elif label == "Employee Name":
                            entries[label] = Entry(addwindow, textvariable=emp_name_var)
                            entries[label].grid(row=row+1, column=col*2+1)
    
            Button(addwindow, text="Add", command=insert_data).grid(row=13,column=1)
    
            addwindow.mainloop()
    
        def export():
            filedir = filedialog.askdirectory()
            filename=filedir+"/Exported file.xlsx"
            sql = pd.read_sql("SELECT * FROM table1", conn)
            sql.to_excel(filename, index=False)
            T.delete("1.0", "end")
            T.insert("1.0", filename)
    
            
    
        def add_emp():
            j = Toplevel()
            j.title("Add Employee")
    
            Label(j, text="Product ID").grid(row=0, column=0,padx=5,pady=5)
            prod_id_entry = Entry(j)
            prod_id_entry.grid(row=0, column=1,padx=5,pady=5)
    
            Label(j, text="Product Name").grid(row=1, column=0,padx=5,pady=5)
            prod_name_entry = Entry(j)
            prod_name_entry.grid(row=1, column=1,padx=5,pady=5)
    
            Label(j, text="Employee ID").grid(row=2, column=0,padx=5,pady=5)
            emp_id_entry = Entry(j)
            emp_id_entry.grid(row=2, column=1,padx=5,pady=5)
    
            Label(j, text="Employee Name").grid(row=3, column=0,padx=5,pady=5)
            emp_name_entry = Entry(j)
            emp_name_entry.grid(row=3, column=1,padx=5,pady=5)
            
    
            def save_employee_product():
                prod_id = prod_id_entry.get()
                prod_name = prod_name_entry.get()
                emp_id = emp_id_entry.get()
                emp_name = emp_name_entry.get()
                
                new_data = pd.DataFrame({
                "Product ID": [prod_id],
                "Product Name": [prod_name],
                "Employee ID": [emp_id],
                "Employee Name": [emp_name]
                })
                    
                global df
                df = pd.concat([df, new_data], ignore_index=True)
    
                update_option_menus()
                T.delete("1.0", "end")
                T.insert("1.0", "Employee/Product Added Successfully!!!!!!!!!!!!!!!!!")
                j.destroy()
            
            Button(j, text="Save", command=save_employee_product).grid(
                row=4, column=0, columnspan=2,padx=5,pady=5)
            
            j.mainloop()
    
        def update_option_menus():# add with new data in dropdown
            global box
            if box:
                box.destroy()
            selected_option = cont.get()
            if selected_option in ["Product Name", "Product ID", "Employee Name", "Employee ID"]:
                sub_options = df[selected_option].dropna().unique()
                box = OptionMenu(r, sub_cont, *sub_options)
                box.place(x=5, y=90)
    
        def add_user():
            add_user=Toplevel()
            Label(add_user,text="Username").grid(row=1,column=1)
            uname_entry=Entry(add_user)
            Label(add_user,text="Password").grid(row=1,column=3)
            upass_entry=Entry(add_user)
            uname_entry.grid(row=2,column=1,padx=10)
            upass_entry.grid(row=2,column=3,padx=10)
            def append_dic():
                nonlocal uname_entry
                nonlocal upass_entry
                uname = uname_entry.get()
                upass = upass_entry.get()
                if uname in user_pass:
                     err.config(text="User already exist")
                else:
                    #to dic
                    user_pass[uname] = upass
                    # Append data to random.xlsx
                    new_data = pd.DataFrame({"Username": [uname], "Password": [upass]})
                    updated_pass_file = pd.concat([pass_file, new_data], ignore_index=True)
                    updated_pass_file.to_excel("password.xlsx", index=False)
                    T.delete("1.0", "end")
                    T.insert("1.0", "User Added Successfully!!!!!!!!!!!!!!!!!")
                    add_user.destroy()
                    
            Button(add_user,text="Add",command=append_dic).grid(row=4,column=2,padx=10,pady=10)
            err=Label(add_user,text='')
            err.grid(row=5,column=2)
            add_user.mainloop()
        
    
        lb1 = Label(r, text="Please select the options")
        options = df.columns
        cont.set("Select")
        dropdown = OptionMenu(r, cont, *options, command=on_select)
    
        submit_btn = Button(r, text="Submit", command=retrieve)
        if username =="admin":
            
            Label(text="To Add new data:").place(x=250,y=10)
            Button(r, text="Add Data", command=add_data).place(x=253, y=33)
            
            Label(text="To Export:").place(x=1290,y=10)
            Button(r, text="Export", command=export).place(x=1305, y=33)
            
            Label(text="To Add User:").place(x=500,y=10)
            Button(r, text="Add User", command=add_user).place(x=506, y=33)
            
            Label(text="To Add Product/Employee:").place(x=250,y=100)
            Button(r, text="Add", command=add_emp).place(x=253, y=125,height=25,width=60)
    
        elif username in operators:
            Label(text="To Add new data:").place(x=250,y=10)
            Button(r, text="Add", command=add_data).place(x=253, y=33)
    
    
        lb1.place(x=5, y=10)
        dropdown.place(x=5, y=33)
        Label(r, text="Seeking for:").place(x=5, y=70)
        submit_btn.place(x=5, y=125)
        
        def switch_user():
            r.destroy()
            application()
        Button(r, text="Sign out", command=switch_user).place(x=1290, y=125)
    
        T = Text(r,width=190,height=180)
        T.place(x=0, y=169)
        
    
        r.title("Info Retrieval")
        r.state("zoomed")
        r.mainloop()
    
    
    # Login Window
    s = tk.Tk()
    Label(s, text="Username").place(x=55,y=80)
    Label(s, text="Password").place(x=55,y=120)
    
    username_entry = Entry(s)
    password_entry = Entry(s, show="*")
    
    username_entry.place(x=125,y=80)
    password_entry.place(x=125,y=120)
    
    login_button = Button(s, text="Login", command=login)
    login_button.place(x=140,y=160)
    
    error_label = Label(s, text="")
    error_label.place(x=80,y=200)
    
    s.title("Login")
    s.geometry("300x300+540+200")
    s.focus_force()
    s.mainloop()
  
application()