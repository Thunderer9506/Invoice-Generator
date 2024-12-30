import tkinter as tk
from tkinter import ttk

#create a window
window = tk.Tk()
window.title("Tkinter Variables")
window.geometry('600x500')

# ComboBox
items = ('Ice Cream','Pizza','Burger')
food_string = tk.StringVar(value=items[0])
combo = ttk.Combobox(window,values=items # Set the values of combobox
	,textvariable=food_string)
combo.pack()

# Events
combo_label = ttk.Label(window,text='A label')
combo_label.pack()

combo.bind('<<ComboboxSelected>>',lambda event:combo_label.config(text=f"Selected value:{food_string.get()}"))

# SpinBox
spin_Int = tk.IntVar(value=12) # Value will set the value that should be displayed first
spin = ttk.Spinbox(
    window,
    from_=3, # Values will range from '3'
    to=20, # to '20'
    command=lambda:print("a arrow was pressed"),
    textvariable=spin_Int)
spin.bind("<<Increment>>",lambda event:print("up"))
spin.bind("<<Decrement>>",lambda event:print("down")) 
spin.pack()

#run
window.mainloop()