import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk
import os
import openpyxl

window = ctk.CTk()
window.configure(fg_color=("#2A2A2A"))
window.title("Inventory by Louelle")
window.geometry("300x400")
frame = ctk.CTkFrame(window)
frame.configure(fg_color=("#2A2A2A"))
frame.pack()




def enter_data():
    
    date = date_input.get()
    name = name_input.get()
    device = device_input.get()
    issue = typeR_input.get()
    warranty = warranty_input.get()
    Tprice = price_input.get()
    partsP = partsP_input.get()
    
    
    if enter_data:
        
        if date_input.get() and name_input and device_input and typeR_input and warranty_input and price_input and partsP_input:
            price = int(price_input.get())
            parts = int(partsP_input.get())
            total = price - parts
            
            
            messagebox.showinfo(title ="Notice", message="Data Saved")
            
        else:
            messagebox.showwarning(message="Please fill in The blanks")
            
        

   
            
            
            
        filepath = "C:/Users/Public/Documents/data.xlsx"
 
 
        if not os.path.exists(filepath):
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            headings = ["Date" , "Name" , "Device" , "Type of Repair" ,"Warranty" ,"Price" ,"PartsPrice" , "Revenue"]
            sheet.append(headings)
            workbook.save(filepath)
        workbook = openpyxl.load_workbook(filepath)
        sheet = workbook.active
        sheet.append([date_input.get(), name_input.get(), device_input.get(), typeR_input.get() ,warranty_input.get(),price_input.get(),partsP_input.get(),total])
        workbook.save(filepath)
         
        date_input.delete(0 , ctk.END)
        name_input.delete(0 , ctk.END)
        device_input.delete(0 , ctk.END)
        typeR_input.delete(0 , ctk.END)
        warranty_input.delete(0 , ctk.END)
        price_input.delete(0 , ctk.END)
        partsP_input.delete(0 , ctk.END)

headings = ctk.CTkLabel(frame,text= "INVENTORY SYSTEM" , font = ctk.CTkFont(size = 25, weight="bold"), text_color= "white" )
headings.grid(row= 0 ,column= 0 ,padx= 5 , pady=0)

client_data = ctk.CTkFrame(frame )
client_data.grid(row =1 , column= 0 ,padx= 10 , pady=5 )
client_data.configure(fg_color= ("#434242"))

date_label = ctk.CTkLabel(client_data, text= "Date:" ,font = ctk.CTkFont(size = 15, weight="bold"), text_color= "white" )
date_label.grid(row =0 , column = 0)

date_input = ctk.CTkEntry(client_data , placeholder_text = "enter date" )
date_input.grid(row =0 , column= 1,padx= 10 , pady=10 )

name_label = ctk.CTkLabel(client_data, text= "Client Name:", font = ctk.CTkFont(size = 15, weight="bold"), text_color= "white" )
name_label.grid(row =1 , column = 0)

name_input = ctk.CTkEntry(client_data , placeholder_text = "enter name")
name_input.grid(row =1 , column= 1,padx= 10 , pady=10 )

device_label = ctk.CTkLabel(client_data, text= "Device:", font = ctk.CTkFont(size = 15, weight="bold"), text_color= "white" )
device_label.grid(row =2 , column = 0)

device_input = ctk.CTkEntry(client_data , placeholder_text = "enter device")
device_input.grid(row =2 , column= 1,padx= 10 , pady=10 )

typeR_label = ctk.CTkLabel(client_data , text="Type of Repair:", font = ctk.CTkFont(size = 15, weight="bold"), text_color= "white" )
typeR_label.grid(row =3 , column= 0)

typeR_input = ctk.CTkEntry(client_data , placeholder_text = "enter type of repair")
typeR_input.grid(row =3 , column= 1 ,padx= 10 , pady=10 )

warranty_label = ctk.CTkLabel(client_data , text="Warranty:", font = ctk.CTkFont(size = 15, weight="bold"), text_color= "white" )
warranty_label.grid(row =4 , column= 0)

warranty_input = tk.Spinbox(client_data ,from_=0 , to="infinity")
warranty_input.grid(row =4 , column= 1 ,padx= 10 , pady=10 )

price_label = ctk.CTkLabel(client_data , text="Price of Repair:", font = ctk.CTkFont(size = 15, weight="bold"), text_color= "white" )
price_label.grid(row =5 , column= 0)

price_input = tk.Spinbox(client_data ,from_= 0 , to="infinity")
price_input.grid(row =5 , column= 1,padx= 10 , pady=10 )

partsP_label = ctk.CTkLabel(client_data , text="Parts Price:", font = ctk.CTkFont(size = 15, weight="bold"), text_color= "white" )
partsP_label.grid(row =6 , column= 0)

partsP_input = tk.Spinbox(client_data ,from_= 0 , to=10000)
partsP_input.grid(row =6 , column= 1,padx= 10 , pady=10 )

confirm_button = ctk.CTkButton(client_data, text= "Save Data" , command= enter_data , font = ctk.CTkFont(size = 15, weight="bold"))
confirm_button.grid(row =7 , column= 1 , sticky= "news" ,padx= 10 , pady=10 )




window.mainloop()