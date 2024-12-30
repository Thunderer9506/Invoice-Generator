import os
import subprocess
from datetime import datetime
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
import docx

class InvoiceAutomation:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Invoice Automation")
        self.root.geometry("600x600")

        self.upperFrame = ttk.Frame(self.root)
        self.upperFrame.place(x = 0, y = 0)
        self.upperFrame.columnconfigure((0,1,2,3), weight=1)
        self.upperFrame.rowconfigure((0,1,2), weight=1)

        self.midFrame = ttk.Frame(self.root)
        self.midFrame.place(relx = 0, rely = 1)
        self.midFrame.columnconfigure((0,1,2,3), weight=1)
        self.midFrame.rowconfigure((0,1,2), weight=1)

        self.endFrame = ttk.Frame(self.root)
        self.endFrame.place(relx = 0, rely = 2)
        self.endFrame.columnconfigure((0,1,2,3), weight=1)
        self.endFrame.rowconfigure((0,1,2), weight=1)

        self.date_label = ttk.Label(self.upperFrame, text="Date")
        self.invoice_label = ttk.Label(self.upperFrame, text="Invoice Number")

        self.clientName_label = ttk.Label(self.upperFrame, text="Client Name")
        self.clientAddress_label = ttk.Label(self.upperFrame, text="Client Address")
        self.clientGST_label = ttk.Label(self.upperFrame, text="Client GST")
        
        # self.description1_label = tk.Label(self.root, text="Description1")
        # self.description2_label = tk.Label(self.root, text="Description2")
        # self.description3_label = tk.Label(self.root, text="Description3")
        # self.description4_label = tk.Label(self.root, text="Description4")
        # self.description5_label = tk.Label(self.root, text="Description5")

        # self.quantity1_label = tk.Label(self.root, text="Quantity1")
        # self.quantity2_label = tk.Label(self.root, text="Quantity2")
        # self.quantity3_label = tk.Label(self.root, text="Quantity3")
        # self.quantity4_label = tk.Label(self.root, text="Quantity4")
        # self.quantity5_label = tk.Label(self.root, text="Quantity5")

        # self.rate1_label = tk.Label(self.root, text="Rate1")
        # self.rate2_label = tk.Label(self.root, text="Rate2")
        # self.rate3_label = tk.Label(self.root, text="Rate3")
        # self.rate4_label = tk.Label(self.root, text="Rate4")
        # self.rate5_label = tk.Label(self.root, text="Rate5")

        # self.amount1_label = tk.Label(self.root, text="Amount1")
        # self.amount2_label = tk.Label(self.root, text="Amount2")
        # self.amount3_label = tk.Label(self.root, text="Amount3")
        # self.amount4_label = tk.Label(self.root, text="Amount4")
        # self.amount5_label = tk.Label(self.root, text="Amount5")

        # self.subtotal_label = tk.Label(self.root, text="Subtotal")
        # self.igst_label = tk.Label(self.root, text="IGST")
        # self.sgst_label = tk.Label(self.root, text="SGST")
        # self.cgst_label = tk.Label(self.root, text="CGST")
        # self.total_label = tk.Label(self.root, text="Total")
        # self.totalInword_label = tk.Label(self.root, text="Total In Word")

        #Create entry widget for every label 
        self.date_entry = ttk.Entry(self.upperFrame,width=30)
        self.invoice_entry = ttk.Entry(self.upperFrame,width=30)

        self.clientName_entry = ttk.Entry(self.upperFrame,width=30)
        self.clientAddress_entry = ttk.Entry(self.upperFrame,width=30)
        self.clientGST_entry = ttk.Entry(self.upperFrame,width=30)

        # self.description1_entry = tk.Entry(self.root)
        # self.description2_entry = tk.Entry(self.root)
        # self.description3_entry = tk.Entry(self.root)
        # self.description4_entry = tk.Entry(self.root)
        # self.description5_entry = tk.Entry(self.root)

        # self.quantity1_entry = tk.Entry(self.root)
        # self.quantity2_entry = tk.Entry(self.root)
        # self.quantity3_entry = tk.Entry(self.root)
        # self.quantity4_entry = tk.Entry(self.root)
        # self.quantity5_entry = tk.Entry(self.root)

        # self.rate1_entry = tk.Entry(self.root)
        # self.rate2_entry = tk.Entry(self.root)
        # self.rate3_entry = tk.Entry(self.root)
        # self.rate4_entry = tk.Entry(self.root)
        # self.rate5_entry = tk.Entry(self.root)

        # self.amount1_entry = tk.Entry(self.root)
        # self.amount2_entry = tk.Entry(self.root)
        # self.amount3_entry = tk.Entry(self.root)
        # self.amount4_entry = tk.Entry(self.root)
        # self.amount5_entry = tk.Entry(self.root)

        # self.subtotal_entry = tk.Entry(self.root)
        # self.igst_entry = tk.Entry(self.root)
        # self.sgst_entry = tk.Entry(self.root)
        # self.cgst_entry = tk.Entry(self.root)
        # self.total_entry = tk.Entry(self.root)
        # self.totalInword_entry = tk.Entry(self.root)

        # self.create_Button = tk.Button(self.root, text="Create Invoice", command=self.create_invoice)

        # Now pack every widget
        # Create a canvas and a scrollbar
        

        # Now use grid to place widgets in the scrollable frame
        self.date_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.date_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        self.invoice_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.invoice_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        self.clientName_label.grid(row=0, column=2, padx=5, pady=5, sticky="w")
        self.clientName_entry.grid(row=0, column=3, padx=5, pady=5, sticky="ew")

        self.clientAddress_label.grid(row=1, column=2, padx=5, pady=5, sticky="w")
        self.clientAddress_entry.grid(row=1, column=3, padx=5, pady=5, sticky="ew")

        self.clientGST_label.grid(row=2, column=2, padx=5, pady=5, sticky="w")
        self.clientGST_entry.grid(row=2, column=3, padx=5, pady=5, sticky="ew")

        # self.description1_label.grid(row=5, column=0, padx=5, pady=5, sticky="w")
        # self.description1_entry.grid(row=5, column=1, padx=5, pady=5, sticky="ew")

        # self.description2_label.grid(row=6, column=0, padx=5, pady=5, sticky="w")
        # self.description2_entry.grid(row=6, column=1, padx=5, pady=5, sticky="ew")

        # self.description3_label.grid(row=7, column=0, padx=5, pady=5, sticky="w")
        # self.description3_entry.grid(row=7, column=1, padx=5, pady=5, sticky="ew")

        # self.description4_label.grid(row=8, column=0, padx=5, pady=5, sticky="w")
        # self.description4_entry.grid(row=8, column=1, padx=5, pady=5, sticky="ew")

        # self.description5_label.grid(row=9, column=0, padx=5, pady=5, sticky="w")
        # self.description5_entry.grid(row=9, column=1, padx=5, pady=5, sticky="ew")

        # self.quantity1_label.grid(row=10, column=0, padx=5, pady=5, sticky="w")
        # self.quantity1_entry.grid(row=10, column=1, padx=5, pady=5, sticky="ew")

        # self.quantity2_label.grid(row=11, column=0, padx=5, pady=5, sticky="w")
        # self.quantity2_entry.grid(row=11, column=1, padx=5, pady=5, sticky="ew")

        # self.quantity3_label.grid(row=12, column=0, padx=5, pady=5, sticky="w")
        # self.quantity3_entry.grid(row=12, column=1, padx=5, pady=5, sticky="ew")

        # self.quantity4_label.grid(row=13, column=0, padx=5, pady=5, sticky="w")
        # self.quantity4_entry.grid(row=13, column=1, padx=5, pady=5, sticky="ew")

        # self.quantity5_label.grid(row=14, column=0, padx=5, pady=5, sticky="w")
        # self.quantity5_entry.grid(row=14, column=1, padx=5, pady=5, sticky="ew")

        # self.rate1_label.grid(row=15, column=0, padx=5, pady=5, sticky="w")
        # self.rate1_entry.grid(row=15, column=1, padx=5, pady=5, sticky="ew")

        # self.rate2_label.grid(row=16, column=0, padx=5, pady=5, sticky="w")
        # self.rate2_entry.grid(row=16, column=1, padx=5, pady=5, sticky="ew")

        # self.rate3_label.grid(row=17, column=0, padx=5, pady=5, sticky="w")
        # self.rate3_entry.grid(row=17, column=1, padx=5, pady=5, sticky="ew")

        # self.rate4_label.grid(row=18, column=0, padx=5, pady=5, sticky="w")
        # self.rate4_entry.grid(row=18, column=1, padx=5, pady=5, sticky="ew")

        # self.rate5_label.grid(row=19, column=0, padx=5, pady=5, sticky="w")
        # self.rate5_entry.grid(row=19, column=1, padx=5, pady=5, sticky="ew")

        # self.amount1_label.grid(row=20, column=0, padx=5, pady=5, sticky="w")
        # self.amount1_entry.grid(row=20, column=1, padx=5, pady=5, sticky="ew")

        # self.amount2_label.grid(row=21, column=0, padx=5, pady=5, sticky="w")
        # self.amount2_entry.grid(row=21, column=1, padx=5, pady=5, sticky="ew")

        # self.amount3_label.grid(row=22, column=0, padx=5, pady=5, sticky="w")
        # self.amount3_entry.grid(row=22, column=1, padx=5, pady=5, sticky="ew")

        # self.amount4_label.grid(row=23, column=0, padx=5, pady=5, sticky="w")
        # self.amount4_entry.grid(row=23, column=1, padx=5, pady=5, sticky="ew")

        # self.amount5_label.grid(row=24, column=0, padx=5, pady=5, sticky="w")
        # self.amount5_entry.grid(row=24, column=1, padx=5, pady=5, sticky="ew")

        # self.subtotal_label.grid(row=25, column=0, padx=5, pady=5, sticky="w")
        # self.subtotal_entry.grid(row=25, column=1, padx=5, pady=5, sticky="ew")

        # self.igst_label.grid(row=26, column=0, padx=5, pady=5, sticky="w")
        # self.igst_entry.grid(row=26, column=1, padx=5, pady=5, sticky="ew")

        # self.sgst_label.grid(row=27, column=0, padx=5, pady=5, sticky="w")
        # self.sgst_entry.grid(row=27, column=1, padx=5, pady=5, sticky="ew")

        # self.cgst_label.grid(row=28, column=0, padx=5, pady=5, sticky="w")
        # self.cgst_entry.grid(row=28, column=1, padx=5, pady=5, sticky="ew")

        # self.total_label.grid(row=29, column=0, padx=5, pady=5, sticky="w")
        # self.total_entry.grid(row=29, column=1, padx=5, pady=5, sticky="ew")

        # self.totalInword_label.grid(row=30, column=0, padx=5, pady=5, sticky="w")
        # self.totalInword_entry.grid(row=30, column=1, padx=5, pady=5, sticky="ew")

        # self.create_Button.grid(row=31, column=0, columnspan=2, padx=5, pady=5, sticky="ew")

        self.root.mainloop()

    @staticmethod
    def replace_text(paragraph, old_text, new_text):
        if old_text in paragraph.text:
            paragraph.text = paragraph.text.replace(old_text,new_text)

    def create_invoice(self):
        doc = docx.Document("template.docx")
        try:
            replacements = {
                "[Date]": self.date_entry.get(),
                "[Invoice]": self.invoice_entry.get(),

                "[clientName]" : self.clientName_entry.get().title(),
                "[clientAddress]" : self.clientAddress_entry.get(),
                "[clientGST]" : self.clientGST_entry.get(),
            }
        except ValueError:
            messagebox.showerror(title='Error',message="Invalid amount or price")
            return
        
        for paragraph in list(doc.paragraphs):
            for old_text, new_text in replacements.items():
                self.replace_text(paragraph,old_text,new_text)
        
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for old_text, new_text in replacements.items():
                            self.replace_text(paragraph,old_text,new_text)
        save_path = filedialog.asksaveasfilename(defaultextension=".pdf",filetypes=[('PDF documents','*.pdf')])
        doc.save('filled.docx')

        try:
            subprocess.run(args=[r'C:\Libre Office\program\soffice.exe', '--headless', '--convert-to', 'pdf', 'filled.docx', '--outdir', '.'], check=True)
            os.rename(src="filled.pdf", dst=save_path)
            messagebox.showinfo(title="success", message="Invoice created and saved successfully")
        except subprocess.CalledProcessError as e:
            messagebox.showerror(title="Error", message=f"Failed to convert DOCX to PDF: {e}")
        except PermissionError as e:
            messagebox.showerror(title="Error", message=f"Permission denied: {e}")


if __name__ == "__main__":
    InvoiceAutomation()