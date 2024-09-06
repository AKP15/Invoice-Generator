import tkinter as tk
from tkinter import ttk
from docxtpl import DocxTemplate
import datetime
from tkinter import messagebox

window=tk.Tk()
window.title('Invoice Generator')
window.geometry('800x500')
window.minsize(800,500)

class TopWidget(ttk.Frame):
        def __init__(self,parent,label_text):
                super().__init__(master=parent)
                self.rowconfigure((0,1),weight=1)
                self.columnconfigure((0),weight=1,uniform='a')
                ttk.Label(self,text=label_text).grid(row=0,column=0,sticky='nesw',pady=10)
                self.entry=ttk.Entry(self)
                self.entry.grid(row=1,column=0,sticky='nesw')
                #self.pack(expand=True,fill='both',padx=10,pady=10)

        def get_input(self):
                x=self.entry.get()
                self.entry.delete(0,tk.END)
                return x
        
        def clear(self):
                return self.entry.delete(0,tk.END)
                
class Topspin(ttk.Frame):
        def __init__(self,parent,label_text):
                super().__init__(master=parent)
                self.rowconfigure((0,1),weight=1)
                self.columnconfigure((0),weight=1,uniform='a')
                ttk.Label(self,text=label_text).grid(row=0,column=0,sticky='nesw',pady=5)
                self.spin=ttk.Spinbox(self,from_= 1,to= 4000)
                self.spin.grid(row=1,column=0,sticky='nesw')

        def get_input(self):
                x=self.spin.get()
                self.spin.delete(0,tk.END)
                return x
        
        def clear(self):
                return self.spin.delete(0,tk.END)
                
invoice_list=[]             
def add():
        qty=int(Qty.get_input())
        desc=Description.get_input()
        price=float(Price.get_input())
        line_total=qty*price
        #print(Description.get_input())
        #print(type(Description.get_input()))

        invoice_items=[qty,desc,price,line_total]
        tree.insert('',0,values=invoice_items)
        invoice_list.append(invoice_items)
        

def clearAll():
        firstName.clear()
        secondName.clear()
        Phone.clear()
        Qty.clear()
        Price.clear()
        Description.clear()
        tree.delete(*tree.get_children())
        invoice_list.clear()

def generate():
        doc=DocxTemplate('invoice_template.docx')
        name= firstName.get_input() + secondName.get_input()
        phone=Phone.get_input()
        subtotal=sum(item[3] for item in invoice_list)
        salestax=0.1
        total=subtotal*(1-salestax)
        doc.render({"name":name, 
            "phone":phone,
            "invoice_list": invoice_list,
            "subtotal":subtotal,
            "salestax":salestax,
            "total":total})
        doc_name = "new_invoice" + name + datetime.datetime.now().strftime("%Y-%m-%d-%H%M%S") + ".docx"
        doc.save(doc_name)
    
        messagebox.showinfo("Invoice Complete", "Invoice Complete")
    
        clearAll()



#main frame
main=ttk.Frame(window)
main.columnconfigure((0,1,2,3),weight=1,uniform='a')
main.rowconfigure((0,1,2,3),weight=1)

topFrame=ttk.Frame(main)
middleFrame=ttk.Frame(main)
buttomFrame=ttk.Frame(main)

#main Frame layout
topFrame.grid(row=0,column=0,sticky='nesw',padx=10,columnspan=4)
#ttk.Label(topFrame,background='red').pack(expand=True,fill="both")
middleFrame.grid(row=1,column=0,sticky='nesw',padx=10,rowspan=2,columnspan=4)
#ttk.Label(topFrame,background='blue').pack(expand=True,fill="both")
buttomFrame.grid(row=3,column=0,sticky='nesw',padx=10,columnspan=4)
#ttk.Label(topFrame,background='green').pack(expand=True,fill="both")
main.pack(expand=True,fill='both')

#top frame widget
topFrame.columnconfigure((0,1,2),weight=1,uniform='a')
topFrame.rowconfigure((0,1,2),weight=1,uniform='a')

firstName=TopWidget(topFrame,'First Name')
firstName.grid(row=0,column=0,sticky='nesw',padx=10)
secondName=TopWidget(topFrame,'Second Name')
secondName.grid(row=0,column=1,sticky='nesw',padx=10)
Phone=TopWidget(topFrame,'Phone')
Phone.grid(row=0,column=2,sticky='nesw',padx=10)
Qty=Topspin(topFrame,'Qty')
Qty.grid(row=1,column=0,sticky='nesw',padx=10)
Description=TopWidget(topFrame,'Description')
Description.grid(row=1,column=1,sticky='nesw',padx=10)
Price=Topspin(topFrame,'Price')
Price.grid(row=1,column=2,sticky='nesw',padx=10)
button=ttk.Button(topFrame,text='Add Item',command=add).grid(row=2,column=2,sticky='ew',padx=5,pady=10)

#middle frame widget
columns=('qty','desc','price','total')
tree=ttk.Treeview(middleFrame,columns=columns,show='headings')
tree.heading('qty',text='Qty')
tree.heading('desc',text='Description')
tree.heading('price',text='Unit Price')
tree.heading('total',text='Total')
tree.pack(expand=True,fill='both',padx=10,pady=5)

#button frame widget
generate_invoice_button=ttk.Button(buttomFrame,text='generate',command=generate)
new_invoice_button=ttk.Button(buttomFrame,text='New',command=clearAll)
generate_invoice_button.pack(side='left',expand=True,fill='x',padx=10)
new_invoice_button.pack(side='left',expand=True,fill='x',padx=10)

window.mainloop()

