#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Test combobox StringVar behavior
"""
import tkinter as tk
from tkinter import ttk

root = tk.Tk()
root.title("Combobox Test")

# Test 1: Create StringVar, then Combobox
var1 = tk.StringVar()
combo1 = ttk.Combobox(root, textvariable=var1, values=['', 'Option1', 'Option2', 'Option3'], state="readonly")
combo1.pack(pady=10)

def on_select1(event):
    print(f"Selected (method 1): var1.get() = '{var1.get()}'")
    print(f"Selected (method 1): combo1.get() = '{combo1.get()}'")

combo1.bind('<<ComboboxSelected>>', on_select1)

# Test 2: Programmatically set value
def set_value():
    print("\nSetting var1 to 'Option2'...")
    var1.set('Option2')
    print(f"After set: var1.get() = '{var1.get()}'")
    print(f"After set: combo1.get() = '{combo1.get()}'")

btn = ttk.Button(root, text="Set to Option2", command=set_value)
btn.pack(pady=10)

# Test 3: Read value on demand
def read_value():
    print(f"\nReading value:")
    print(f"  var1.get() = '{var1.get()}'")
    print(f"  combo1.get() = '{combo1.get()}'")

btn2 = ttk.Button(root, text="Read Value", command=read_value)
btn2.pack(pady=10)

print("Test started. Try selecting from dropdown and clicking buttons.")
root.mainloop()
