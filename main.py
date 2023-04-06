from tkinter import *
from tkinter import ttk
from docxtpl import DocxTemplate
from docx2pdf import convert
import time

invoiceList=[]

def newInvoice():
    firstNameEntry.delete (0, END)
    lastNameEntry.delete (0, END)
    phoneEntry.delete (0, END)
    tree.delete(*tree.get_children())
    clearItem()
    invoiceList.clear()
    
    
def generateInvoice():
    newDoc = DocxTemplate('Invoice Template.docx')
    firstName = firstNameEntry.get()
    lastName = lastNameEntry.get()
    phone = phoneEntry.get()
    total = sum(item[3] for item in invoiceList)
    newDoc.render({
        'name':firstName+' '+lastName,
        'phone':phone,
        'invoice_list':invoiceList,
        'total':total
    })
    newDoc.save('Invoice/Word/'+firstName+lastName+'.docx')
    time.sleep(2.0)
    convert('Invoice/Word/'+firstName+lastName+'.docx',
            'Invoice/'+firstName+lastName+'.pdf')

def clearItem():
    quantitySpinBox.delete(0,END)
    quantitySpinBox.insert (0,'1')
    descriptionEntry.delete (0, END)
    unitPriceSpinBox.delete(0, END)
    unitPriceSpinBox.insert(0, '0')

def addItem():
    qty = int(quantitySpinBox.get())
    desc = descriptionEntry.get()
    price = float(unitPriceSpinBox.get())
    lineTotal = qty*price
    invoiceItem = [qty,desc,price,lineTotal]
    tree.insert('',0,values=invoiceItem)
    invoiceList.append(invoiceItem)
    clearItem()

window = Tk()
# window.geometry('1280x720')
window.title('Invoice Generator')
windowIcon = PhotoImage(file='Assets/img/receipt.png') # Converts into Photo image
window.iconphoto(True,windowIcon)

frame = Frame(window)
frame.pack(padx=20,pady=10)

firstNameLabel = Label(frame,
                       text='First Name',
                       font=('Poppins',16,'bold'))
firstNameLabel.grid(row=0,column=0)
firstNameEntry = Entry(frame,font=('Poppins',16))
firstNameEntry.grid(row=1,column=0)

lastNameLabel = Label(frame,
                       text='Last Name',
                       font=('Poppins',16,'bold'))
lastNameLabel.grid(row=0,column=1)
lastNameEntry = Entry(frame,font=('Poppins',16))
lastNameEntry.grid(row=1,column=1)


phoneLabel = Label(frame,
                       text='Phone Number',
                       font=('Poppins',16,'bold'))
phoneLabel.grid(row=0,column=2)
phoneEntry = Entry(frame,font=('Poppins',16))
phoneEntry.grid(row=1,column=2)

quantityLabel = Label(frame,
                       text='Quantity',
                       font=('Poppins',16,'bold'))
quantityLabel.grid(row=2,column=0)
quantitySpinBox = Spinbox(frame,font=('Poppins',16),from_=1,to=50)
quantitySpinBox.grid(row=3,column=0)

descriptionLabel = Label(frame,
                       text='Description',
                       font=('Poppins',16,'bold'))
descriptionLabel.grid(row=2,column=1)
descriptionEntry = Entry(frame,font=('Poppins',16))
descriptionEntry.grid(row=3,column=1)


unitPriceLabel = Label(frame,
                       text='Unit Price',
                       font=('Poppins',16,'bold'))
unitPriceLabel.grid(row=2,column=2)
unitPriceSpinBox = Spinbox(frame,font=('Poppins',16),from_=0,to_=999999)
unitPriceSpinBox.grid(row=3,column=2)

addItemButton = Button(frame,
                       font=('Poppins',14,'bold'),
                       text='Add Item',
                       fg='#fff',
                       bg='#444',
                       command= addItem)
addItemButton.grid(row=4,column=1,pady=30)

tableColumns = ('qty','desc','price','total')
tree = ttk.Treeview(frame,columns=tableColumns,show='headings')
style=ttk.Style()

style.configure('Treeview',font=('Poppins',14),rowheight=30)
style.configure('Treeview.Heading',font=('Poppins',16,'bold'))

tree.heading('qty',text='Quantity',)
tree.heading('desc',text='Description')
tree.heading('price',text='Unit Price')
tree.heading('total',text='Total')

tree.grid(row=5,column=0,columnspan=3,padx=20,pady=10)

saveInvoiceButton = Button(frame,text='Generate Invoice',font=('Poppins',14,'bold'),fg='#fff',bg='#444',command=generateInvoice)
saveInvoiceButton.grid(row=6,column=0,columnspan=3,padx=20,pady=5,sticky='news')

newInvoiceButton = Button(frame,text='New Invoice',font=('Poppins',14,'bold'),fg='#fff',bg='#444',command= newInvoice)
newInvoiceButton.grid(row=7,column=0,columnspan=3,padx=20,pady=5,sticky='news')

window.mainloop()