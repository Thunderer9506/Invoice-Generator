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