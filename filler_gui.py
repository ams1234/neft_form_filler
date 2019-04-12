from tkinter import *
import excel_filler
from num2words import num2words

root = Tk()

def createdoc():
    global search_name
    excel_filler.search_entry_and_generate_doc(search_name.get())

def saveData():

    global payee
    global account_num
    global ifsc
    global bank
    global address
    global amount
    global amount_words
    global date

    payee_name = payee.get()
    payee_account_num = account_num.get()
    payee_ifsc = ifsc.get()
    payee_bank = bank.get()
    payee_address = address.get()
    payee_amount = amount.get()
    payee_amount_words = amount_words.cget('text')
    payee_date = date.get()
    excel_filler.add_entry(payee_name, payee_account_num, payee_ifsc, payee_bank, payee_address, payee_amount, payee_amount_words, payee_date)
    excel_filler.add_to_json(payee_name, payee_account_num, payee_ifsc, payee_bank, payee_address, payee_amount, payee_amount_words, payee_date)

root.minsize(600,400)

root.title("Bank Form Details")
root.configure(background='light grey')

root.grid_columnconfigure(0, minsize=20)
root.grid_columnconfigure(1, minsize=30)
root.grid_columnconfigure(3, minsize=75)
root.grid_rowconfigure(0, minsize=50)
root.grid_rowconfigure(2, minsize=50)
root.grid_rowconfigure(4, minsize=50)


label_1 = Label(root, text ="Payee :", font = (None, 15))
label_1.grid(row=1,column=2,sticky=W)

payee = Entry(root, width = 30)
payee.grid(row=1, column=4, sticky=W)

label_2 = Label(root, text ="Account Number  :", font = (None, 15))
label_2.grid(row=2,column=2,sticky=W)

account_num = Entry(root, width = 30)
account_num.grid(row=2, column=4, sticky=W)

label_3 = Label(root, text ="IFSC   :", font = (None, 15))
label_3.grid(row=3,column=2,sticky=W)

ifsc = Entry(root, width = 30)
ifsc.grid(row=3, column=4, sticky=W)

label_4 = Label(root, text ="Bank Name  :", font = (None, 15))
label_4.grid(row=4,column=2,sticky=W)

address = Entry(root, width = 30)
address.grid(row=4, column=4, sticky=W)

label_5 = Label(root, text ="Bank Address  :", font = (None, 15))
label_5.grid(row=5,column=2,sticky=W)

amount = Entry(root, width = 30)
amount.grid(row=5, column=4, sticky=W)

label_6 = Label(root, text ="Bank  :", font = (None, 15))
label_6.grid(row=6,column=2,sticky=W)

bank = Entry(root, width = 30)
bank.grid(row=6, column=4, sticky=W)

label_7 = Label(root, text ="Amount  :", font = (None, 15))
label_7.grid(row=7,column=2,sticky=W)

amount = Spinbox(root, width = 30)
amount.grid(row=7, column=4, sticky=W)


label_8 = Label(root, text ="Amount in words  :", font = (None, 15))
label_8.grid(row=8,column=2,sticky=W)


amount_words = Label(root,text=num2words(0), font = (None, 15))
amount_words.grid(row=8, column=4, sticky=W)
# amount_words = Entry(root, show="*", width = 30)
# amount_words.grid(row=7, column=4, sticky=W)

label_9 = Label(root,text = "Date  :", font = (None, 15))
label_9.grid(row=9,column=2,sticky=W)

date = Entry(root, width = 30)
date.grid(row=9, column=4, sticky=W)

button = Button(root,text='Save Data',command=saveData, height = 1, width = 10,font = (None, 15))
button.grid(row=10, column=3, sticky=W)

label_10 = Label(root,text = "Neft Form To be Created for:", font = (None, 15))
label_10.grid(row=13,column=2,sticky=W)

search_name = Entry(root, width = 30)
search_name.grid(row=13, column=4, sticky=W)

button = Button(root,text='Print',command=createdoc, height = 1, width = 10,font = (None, 15))
button.grid(row=14, column=3, sticky=W)



root.mainloop()


