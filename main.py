import os
import subprocess
from datetime import datetime
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import docx

class InvoiceAutomation:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Invoice Automation")
        self.root.geometry("500x600")
        
        self.date_label = tk.Label(self.root, text="Date")
        self.invoice_label = tk.Label(self.root, text="Invoice Number")

        self.clientName_label = tk.Label(self.root, text="Client Name")
        self.clientAddress_label = tk.Label(self.root, text="Client Address")
        self.clientGST_label = tk.Label(self.root, text="Client GST")
        
        self.description1_label = tk.Label(self.root, text="Description1")
        self.description2_label = tk.Label(self.root, text="Description2")
        self.description3_label = tk.Label(self.root, text="Description3")
        self.description4_label = tk.Label(self.root, text="Description4")
        self.description5_label = tk.Label(self.root, text="Description5")

        self.quantity1_label = tk.Label(self.root, text="Quantity1")
        self.quantity2_label = tk.Label(self.root, text="Quantity2")
        self.quantity3_label = tk.Label(self.root, text="Quantity3")
        self.quantity4_label = tk.Label(self.root, text="Quantity4")
        self.quantity5_label = tk.Label(self.root, text="Quantity5")

        self.rate1_label = tk.Label(self.root, text="Rate1")
        self.rate2_label = tk.Label(self.root, text="Rate2")
        self.rate3_label = tk.Label(self.root, text="Rate3")
        self.rate4_label = tk.Label(self.root, text="Rate4")
        self.rate5_label = tk.Label(self.root, text="Rate5")

        self.amount1_label = tk.Label(self.root, text="Amount1")
        self.amount2_label = tk.Label(self.root, text="Amount2")
        self.amount3_label = tk.Label(self.root, text="Amount3")
        self.amount4_label = tk.Label(self.root, text="Amount4")
        self.amount5_label = tk.Label(self.root, text="Amount5")

        self.subtotal_label = tk.Label(self.root, text="Subtotal")
        self.igst_label = tk.Label(self.root, text="IGST")
        self.sgst_label = tk.Label(self.root, text="SGST")
        self.cgst_label = tk.Label(self.root, text="CGST")
        self.total_label = tk.Label(self.root, text="Total")
        self.totalInword_label = tk.Label(self.root, text="Total In Word")

        #Create entry widget for every label 
        self.date_entry = tk.Entry(self.root)
        self.invoice_entry = tk.Entry(self.root)

        self.clientName_entry = tk.Entry(self.root)
        self.clientAddress_entry = tk.Entry(self.root)
        self.clientGST_entry = tk.Entry(self.root)

        self.description1_entry = tk.Entry(self.root)
        self.description2_entry = tk.Entry(self.root)
        self.description3_entry = tk.Entry(self.root)
        self.description4_entry = tk.Entry(self.root)
        self.description5_entry = tk.Entry(self.root)

        self.quantity1_entry = tk.Entry(self.root)
        self.quantity2_entry = tk.Entry(self.root)
        self.quantity3_entry = tk.Entry(self.root)
        self.quantity4_entry = tk.Entry(self.root)
        self.quantity5_entry = tk.Entry(self.root)

        self.rate1_entry = tk.Entry(self.root)
        self.rate2_entry = tk.Entry(self.root)
        self.rate3_entry = tk.Entry(self.root)
        self.rate4_entry = tk.Entry(self.root)
        self.rate5_entry = tk.Entry(self.root)

        self.amount1_entry = tk.Entry(self.root)
        self.amount2_entry = tk.Entry(self.root)
        self.amount3_entry = tk.Entry(self.root)
        self.amount4_entry = tk.Entry(self.root)
        self.amount5_entry = tk.Entry(self.root)

        self.subtotal_entry = tk.Entry(self.root)
        self.igst_entry = tk.Entry(self.root)
        self.sgst_entry = tk.Entry(self.root)
        self.cgst_entry = tk.Entry(self.root)
        self.total_entry = tk.Entry(self.root)
        self.totalInword_entry = tk.Entry(self.root)

        self.create_Button = tk.Button(self.root, text="Create Invoice", command=self.create_invoice)

        padding_options = {'fill': 'x', 'expand':True, 'padx': 5, 'pady': 5} 

        # Now pack every widget
        self.date_label.pack(padding_options)
        self.date_entry.pack(padding_options)
        
        self.invoice_label.pack(padding_options)
        self.invoice_entry.pack(padding_options)
        
        self.clientName_label.pack(padding_options)
        self.clientName_entry.pack(padding_options)
        
        self.clientAddress_label.pack(padding_options)
        self.clientAddress_entry.pack(padding_options)
        
        self.clientGST_label.pack(padding_options)
        self.clientGST_entry.pack(padding_options)
        
        self.description1_label.pack(padding_options)
        self.description1_entry.pack(padding_options)
        
        self.description2_label.pack(padding_options)
        self.description2_entry.pack(padding_options)
        
        self.description3_label.pack(padding_options)
        self.description3_entry.pack(padding_options)
        
        self.description4_label.pack(padding_options)
        self.description4_entry.pack(padding_options)
        
        self.description5_label.pack(padding_options)
        self.description5_entry.pack(padding_options)
        
        self.quantity1_label.pack(padding_options)
        self.quantity1_entry.pack(padding_options)
        
        self.quantity2_label.pack(padding_options)
        self.quantity2_entry.pack(padding_options)
        
        self.quantity3_label.pack(padding_options)
        self.quantity3_entry.pack(padding_options)
        
        self.quantity4_label.pack(padding_options)
        self.quantity4_entry.pack(padding_options)
        
        self.quantity5_label.pack(padding_options)
        self.quantity5_entry.pack(padding_options)
        
        self.rate1_label.pack(padding_options)
        self.rate1_entry.pack(padding_options)
        
        self.rate2_label.pack(padding_options)
        self.rate2_entry.pack(padding_options)
        
        self.rate3_label.pack(padding_options)
        self.rate3_entry.pack(padding_options)
        
        self.rate4_label.pack(padding_options)
        self.rate4_entry.pack(padding_options)
        
        self.rate5_label.pack(padding_options)
        self.rate5_entry.pack(padding_options)
        
        self.amount1_label.pack(padding_options)
        self.amount1_entry.pack(padding_options)
        
        self.amount2_label.pack(padding_options)
        self.amount2_entry.pack(padding_options)
        
        self.amount3_label.pack(padding_options)
        self.amount3_entry.pack(padding_options)
        
        self.amount4_label.pack(padding_options)
        self.amount4_entry.pack(padding_options)
        
        self.amount5_label.pack(padding_options)
        self.amount5_entry.pack(padding_options)
        
        self.subtotal_label.pack(padding_options)
        self.subtotal_entry.pack(padding_options)
        
        self.igst_label.pack(padding_options)
        self.igst_entry.pack(padding_options)
        
        self.sgst_label.pack(padding_options)
        self.sgst_entry.pack(padding_options)
        
        self.cgst_label.pack(padding_options)
        self.cgst_entry.pack(padding_options)
        
        self.total_label.pack(padding_options)
        self.total_entry.pack(padding_options)
        
        self.totalInword_label.pack(padding_options)
        self.totalInword_entry.pack(padding_options)
        
        self.create_Button.pack(padding_options)

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