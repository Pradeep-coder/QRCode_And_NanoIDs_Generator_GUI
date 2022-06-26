from tkinter import *
import pyqrcode
from pyqrcode import QRCode
import png
import xlrd
from time import sleep
import random
from tkinter import messagebox
from tkinter import filedialog
from nanoid import generate
import re
import csv

random_filename_1 = '0123456789'
random_filename_2 = '0123456789'
symbols = random_filename_1 + random_filename_2
length = 4
file_name = ''.join(random.sample(symbols, length))

class main_class():
    def __init__(self):
        self.window = Tk()
        self.window.title("QR Code Generator")
        self.window.geometry("1200x300")
        self.window.iconbitmap(r'.\qr_logo_i.ico')

    def main_func(self):
        self.l1 = Label(self.window, text="QR CODE GENERATOR",font=('helvetica', 12, 'bold'))
        self.l1.grid(row=0, column=1, columnspan=6)

        self.l2 = Label(self.window, text="Enter the URL: ")
        self.l2.grid(row=3, column=1, columnspan=1)
        self.e1 = Entry(self.window, width=45)
        self.e1.grid(row=3, column=3, columnspan=2)

        self.tkvar = StringVar(self.window)
        choices = {5, 6, 7, 8, 9, 10}
        self.tkvar.set(10)

        self.l3 = Label(self.window, text="Scale the QR Codes to: ")
        self.l3.grid(row=3, column=6, columnspan=2)
        self.dropdown_scale = OptionMenu(self.window, self.tkvar, *choices)
        self.dropdown_scale.grid(row=3, column=10, columnspan=1)

        self.l4 = Label(self.window, text="Want to generate multiple QR codes?",font=('helvetica', 10, 'bold'))
        self.l4.grid(row=15, rowspan=1, column=3)
        self.l5 = Label(self.window, text="")
        self.l5.grid(row=19, rowspan=1, column=3)

        self.l6 = Label(self.window, text="Want to generate Nano IDs?",font=('helvetica', 10, 'bold'))
        self.l6.grid(row=15, rowspan=1, column=11)

        self.v1 = DoubleVar()
        self.s1 = Scale(self.window, variable=self.v1, from_=1, to=100, orient=HORIZONTAL)
        self.s1.grid(row=17, column=11)

        self.e2 = Entry(self.window, width=25)
        self.e2.grid(row=21, rowspan=1, column=11)

        self.b6 = Button(self.window, text="Generate Nano IDs",command=self.create_multisub_nanoids)
        self.b6.grid(row=25, rowspan=1, column=11)
        self.b5 = Label(self.window, text="Enter output CSV file name")
        self.b5.grid(row=23, rowspan=1, column=11)
        self.b4 = Label(self.window, text="Choose range above")
        self.b4.grid(row=19, rowspan=1, column=11)

        self.b3 = Button(self.window, text="Generate Multi QR", command=self.create_multi_qr_sub)
        self.b3.grid(row=21, column=3)
        self.b2 = Button(self.window, text="Choose an excel(.xlsx) file", command=self.create_multi_qr)
        self.b2.grid(row=17, rowspan=1, column=3)
        self.l7 = Label(self.window, text="Note!: In column-A place redirect url ex: https://www.google.com/\n column-B name of the image ex: myqrimg")
        self.l7.grid(row=18, rowspan=1, column=3)

        self.b1 = Button(self.window, text='Generate QR', command=self.create_qr)
        self.b1.grid(row=9, column=1, columnspan=5)

        self.window.mainloop()

    def create_multisub_nanoids(self):
        if not self.e2.get():
            messagebox.showerror(
                "Error", "Enter valid csv file name")
        else:
            if '.csv' not in self.e2.get():
                csv_file = open(self.e2.get() + '.csv', 'w')
                csv_writer = csv.writer(csv_file)
                csv_writer.writerow(['A', 'B', 'RedirectURL(A+B)'])
                for result in range(0, int(self.v1.get())):
                    regex = re.compile('[-_]')
                    result = str(generate(size=12))
                    if (regex.search(result) == None):
                        print(result)
                        csv_writer.writerow(["https://www.google.com/",result, "https://www.google.com/"+result])
                        sleep(0.9)
                    else:
                        print(
                            f"Nano id contains special characters: {result}")
                        csv_writer.writerow(["https://www.google.com/",f"Nano id contains special characters: {result}", "https://www.google.com/" + f"Nano id contains special characters: {result}"])
                csv_file.close()
                messagebox.showinfo("Nano ID", "Nano IDs created successfully!")
            else:
                messagebox.showinfo("Nano ID", "Remove the .csv extension\n We make your work simpler!")

    def create_multi_qr(self):
        global pat
        finame = filedialog.askopenfilename(title='Select an excel file', filetype=(
            ("Excel", "*.xlsx"), ("Excel", "*.xls")))

        pat = finame
        self.l5.configure(text=pat)

    def create_multi_qr_sub(self):
        try:
            loc = pat
            wb = xlrd.open_workbook(loc)
            sheet = wb.sheet_by_index(0)
            sheet.cell_value(0, 0)
            for i in range(sheet.nrows):
                url = sheet.cell_value(i, 0)
                main_url = pyqrcode.create(url)
                main_url.png(sheet.cell_value(i, 1) +'.png', scale=self.tkvar.get())
                sleep(0.03)
            messagebox.showinfo(
                "QR Created", "QR codes are successfully generated")
            self.l5['text'] = ""
            self.b3.configure(state=DISABLED)
        except:
            self.l5['text'] = f"Error! Choose an excel(.xlsx) file"

    def create_qr(self):
        if not self.e1.get():
            messagebox.showerror("Error", "Invalid Input")
        else:
            try:
                url = self.e1.get()
                main_url = pyqrcode.create(url)
                img_name = file_name + '_scale' + self.tkvar.get()
                main_url.png(img_name+'.png', scale=self.tkvar.get())
                messagebox.showinfo(
                    "Success", f"QR Code Created with file name {img_name}")
                self.e1.delete(0, 'end')
            except:
                messagebox.showerror("Error", "Oops...Error occurred")


boom = main_class()
boom.main_func()
