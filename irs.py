from tkinter import *
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np
import csv
import os
# Note, for program to export properly, you have to install 'xlsxwriter' module— "pip install xlsxwriter"


root = Tk()
root.title("Tax Computation Program")
root.geometry("1225x725")
root.pack_propagate(False)
root.resizable(0, 0)
root["bg"]="#f9f9f9"
root.iconbitmap("tax.ico")

#global style

style = ttk.Style()
style.theme_use("clam")
style.configure("T.Treeview",
    background="#fff",
    foreground="black",
    rowheight="25",
    fieldbackground="#fff"
    )


style.map("T.Treeview",
    background=[('selected', '#B2B2B2')],
    foreground=[('selected', 'black')]
    )

# Frame to contain everything
top_frame = Frame(root, bd=2)
top_frame.place(width=1175, height=700, relx=0.020, rely=0.01)
top_frame["bg"]="#e2e2e2"

# Frame for Treeview and buttons
sub_frame = Frame(top_frame, bd=2)
sub_frame.place(height=375, width=1125, relx=0.020, rely=0.01)
sub_frame["bg"]="#f9f9f9"

# Frame for the Treeview
tree_frame = Frame(sub_frame)
tree_frame.place(height=285, width=1125, relx=0, rely=0.084)
tree_frame["bg"]="#f9f9f9"

# Label for header above Treeview
sub_frame_label = Label(sub_frame, text="PAYE COMPUTATION", bg="#79c1f1")
sub_frame_label.place(relx=0, rely=0, relwidth=1, height=30)


# Treeview
tree = ttk.Treeview(tree_frame, style="T.Treeview")
tree.place(relheight=1, relwidth=1)

# Scrollbar for Treeview
treescrollx = Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
treescrolly = Scrollbar(tree_frame, orient="vertical", command=tree.yview)
treescrollx.pack(side="bottom", fill=X)
treescrolly.pack(side="right", fill=Y)
tree.configure(xscrollcommand=treescrollx.set, yscrollcommand=treescrolly.set)

# Set Columns for Treeview
tree["column"] = ("s/n", "name", "designation", "tin", "work", "monthly", "annual_gross", "cra", "pen", "nhf", "nhis", "grat", "taxable", "annual_tax", "monthly_tax")

# Format Columns
tree.column("#0", width=0)
tree.column("s/n", anchor=CENTER, width=35, minwidth=25)
tree.column("name", anchor=CENTER, width=250)
tree.column("designation", anchor=CENTER, width=250)
tree.column("tin", anchor=CENTER, width=200)
tree.column("work", anchor=CENTER, width=150)
tree.column("monthly", anchor=CENTER, width=200)
tree.column("annual_gross", anchor=CENTER, width=200)
tree.column("cra", anchor=CENTER, width=200)
tree.column("pen", anchor=CENTER, width=200)
tree.column("nhf", anchor=CENTER, width=200)
tree.column("nhis", anchor=CENTER, width=200)
tree.column("grat", anchor=CENTER, width=200)
tree.column("taxable", anchor=CENTER, width=200)
tree.column("annual_tax", anchor=CENTER, width=200)
tree.column("monthly_tax", anchor=CENTER, width=200)

# Column headings
tree.heading("#0", text="")
tree.heading("s/n", text="S/N", anchor=CENTER)
tree.heading("name", text="NAME", anchor=CENTER)
tree.heading("designation", text="DESIGNATION", anchor=CENTER)
tree.heading("tin", text="T.I.N", anchor=CENTER)
tree.heading("work", text="MONTHS WORKED", anchor=CENTER)
tree.heading("monthly", text="MONTHLY GROSS INCOME", anchor=CENTER)
tree.heading("annual_gross", text="ANNUAL GROSS INCOME", anchor=CENTER)
tree.heading("cra", anchor=CENTER, text="RELIEF ALLOWANCE")
tree.heading("pen", text="PENSION", anchor=CENTER)
tree.heading("nhf", text="NATIONAL HOUSING FUND", anchor=CENTER)
tree.heading("nhis", text="NHIS", anchor=CENTER)
tree.heading("grat", text="GRATUITIES", anchor=CENTER)
tree.heading("taxable", anchor=CENTER, text="TAXABLE INCOME")
tree.heading("annual_tax", anchor=CENTER, text="ANNUAL TAX DUE")
tree.heading("monthly_tax", anchor=CENTER, text="MONTHLY TAX DUE")



global add_btn


# Frame for label and button of form
info_frame = LabelFrame(top_frame, bd=2)
info_frame.place(height=300, width=600, relx=0.020 , rely=0.56)
info_frame.configure(bg="#f9f9f9")

# Label for title of info_frame
title_label = Label(info_frame, text="TAXPAYER'S DATA", bg="#79c1f1")
title_label.place(relwidth=1, relx=0, rely=0, height=30)

# Label widgets for form
org_label = Label(info_frame, text="Organization", bg="#f9f9f9").place(relx=0.1, rely=0.45)
name_label = Label(info_frame, text="Name", bg="#f9f9f9").place(relx=0.1, rely=0.15)
Des_label = Label(info_frame, text="Designation", bg="#f9f9f9").place(relx=0.1, rely=0.25)
tin_label = Label(info_frame, text="T.I.N", bg="#f9f9f9").place(relx=0.1, rely=0.35)
work_label = Label(info_frame, text="Months Worked", bg="#f9f9f9").place(relx=0.1, rely=0.55)
mon_label = Label(info_frame, text="Monthly Gross Income", bg="#f9f9f9").place(relx=0.1, rely=0.65)
# Change background color for the label widgets


# Entry widgets for the form
name_box = Entry(info_frame, width=25)
name_box.place(relx=0.4, rely=0.15)

des_box = Entry(info_frame, width=25)
des_box.place(relx=0.4, rely=0.25)

tin_box = Entry(info_frame, width=25)
tin_box.place(relx=0.4, rely=0.35)

org_box = Entry(info_frame, width=25)
org_box.place(relx=0.4, rely=0.45)

work_box = Entry(info_frame, width=25)
work_box.place(relx=0.4, rely=0.55)

mon_box = Entry(info_frame, width=25)
mon_box.place(relx=0.4, rely=0.65)


# Function to display form of selected checkbox
def add_exemption():

    # Main frame encapsulating the fields for the exemptions
    main_frame.place(relx=0.325, rely=0.19)

    # Variables to get value from checkboxes' variable (IntVar)
    gra_tax = gVar.get()
    pen_tax = pVar.get()
    ns_tax = n_sVar.get()
    nh_tax = n_hVar.get()

    if gra_tax == 1:
        g_frame.pack()
        g_label = Label(g_frame, text="Gratuities", bg="#f9f9f9")
        g_label.place(relx=0.05, rely=0.25)
        g_field.place(relx=0.325, rely=0.25)

    if pen_tax == 1:
        p_frame.pack()
        p_label = Label(p_frame, text="Pension", bg="#f9f9f9")
        p_label.place(relx=0.05, rely=0.25)
        p_field.place(relx=0.325, rely=0.25)

    if ns_tax == 1:
        ns_frame.pack()
        ns_label = Label(ns_frame, text="NHIS", bg="#f9f9f9")
        ns_label.place(relx=0.05, rely=0.25)
        ns_field.place(relx=0.325, rely=0.25)

    if nh_tax == 1:
        nh_frame.pack()
        nh_label = Label(nh_frame, text="NHF", bg="#f9f9f9")
        nh_label.place(relx=0.05, rely=0.25)
        nh_field.place(relx=0.325, rely=0.25)

    exemp_add_btn["state"] = DISABLED


# Function to clear form displayed above
def clear_exemption():

    # To "forget" all the frames in Tax Exemption section. (Remove from screen)
    main_frame.place_forget()
    g_frame.pack_forget()
    p_frame.pack_forget()
    ns_frame.pack_forget()
    nh_frame.pack_forget()

    # Delete redundant data in the Entry boxes
    g_field.delete(0, END)
    p_field.delete(0, END)
    ns_field.delete(0, END)
    nh_field.delete(0, END)

    # To Deselect the checkbuttons
    g.deselect()
    p.deselect()
    n_s.deselect()
    n_h.deselect()

    exemp_add_btn["state"] = DISABLED

# Frame for Tax Exemption section
exemption_frame = Frame(top_frame, bd=2)
exemption_frame.place(height=300, width=500, relx=0.555 , rely=0.56)
exemption_frame.configure(bg="#f9f9f9")

# Label for title of exemption_frame
exemption_label = Label(exemption_frame, text="TAX EXEMPTIONS", bg="#79c1f1")
exemption_label.place(relwidth=1, relx=0, rely=0, height=30)

# Label to display instructions
label = Label(exemption_frame, text="Select Tax Exemption", bg="#f9f9f9", font="20")
label.pack(pady=(32, 0))

# Interger variables to use for checkbutton
gVar = IntVar()
pVar = IntVar()
n_sVar = IntVar()
n_hVar = IntVar()

# Checkboxes
g = Checkbutton(exemption_frame, text="Gratuities", variable=gVar, onvalue=1, offvalue=0, bg="#f9f9f9")
g.place(relx=0.05, rely=0.25)
p = Checkbutton(exemption_frame, text="Pension", variable=pVar, onvalue=1, offvalue=0, bg="#f9f9f9")
p.place(relx=0.05, rely=0.4)
n_s = Checkbutton(exemption_frame, text="NHIS", variable=n_sVar, onvalue=1, offvalue=0, bg="#f9f9f9")
n_s.place(relx=0.05, rely=0.55)
n_h = Checkbutton(exemption_frame, text="NHF", variable=n_hVar, onvalue=1, offvalue=0, bg="#f9f9f9")
n_h.place(relx=0.05, rely=0.7)

# Buttons for the Checkboxes
exemp_add_btn = Button(exemption_frame, text="Add", width=10, command=add_exemption, state=DISABLED)
exemp_add_btn.place(relx=0.25, rely=0.9)
exemp_clear_btn = Button(exemption_frame, text="Clear", width=10, command=clear_exemption)
exemp_clear_btn.place(relx=0.55, rely=0.9)

# All frames for the exemptions
main_frame = Frame(exemption_frame, bg="#f9f9f9", height=200, width=300) # Main container
g_frame = Frame(main_frame, bg="#f2f2f2", height=50, width=300)   # Gratuities frame
p_frame = Frame(main_frame, bg="#f2f2f2", height=50, width=300)    # Pension frame
ns_frame = Frame(main_frame, bg="#f2f2f2", height=50, width=300)    # NHIS frame
nh_frame = Frame(main_frame, bg="#f2f2f2", height=50, width=300)   # NHF frame

g_field = Entry(g_frame)  # Gratuities entry box
p_field = Entry(p_frame)   # Pension entry box
ns_field = Entry(ns_frame)  # NHIS entry box
nh_field = Entry(nh_frame)   # NHF entry box


count = 0   # Variable to count iid in Treeview
serial = 1  # Variable to count S/N (Serial number) in Treeview. Note, it starts from "1"

def error():

    # To validate "tin_box i.e T.I.N" and prompt user if an invalid value of less or greater than 13 characters is provided.
    if len(tin_box.get()) < 13:
        try:
            loophole = float(name_box.get())
        except ValueError:
            messagebox.showwarning("Tax Identification Number", "Please enter a valid value! (i.e 190XXXXX-0001)")
            add_payer.quit()

    if len(tin_box.get()) > 13:
        try:
            loophole = float(name_box.get())
        except ValueError:
            messagebox.showwarning("Tax Identification Number", "Please enter a valid value! (i.e 190XXXXX-0001)")
            add_payer.quit()


    # To grab the value of "mon_box, i.e Monthly Gross" and convert to float, and spit out a messagebox if it fails to convert
    try:
        float(mon_box.get())
    except ValueError:
        messagebox.showwarning("Monthly Gross Income", "Please enter a valid number!")

# Function to add Tax Payer to the Treeview
def add_payer(event):
    global count
    global serial
    global monVar2
    global annVar2
    global cra_Var2
    global taxVar2
    global mon_tax_Var2
    global ann_tax_Var2
    global g_varr2
    global p_varr2
    global ns_varr2
    global nh_varr2

    # Variable to store input from mon_box as a float
    cash = float(mon_box.get())

    error()

    # Create Variable to "get" input from the form
    nameVar = name_box.get().upper()
    desVar = des_box.get().upper()
    tinVar = tin_box.get()
    workVar = work_box.get()
    monVar = cash
    monVar2 = "{:,.2f}".format(monVar) # Separate the figures inputed with commas e.g 1000000 = 1,000,000


    if main_frame.winfo_ismapped():

        # Variables to get value from checkboxes' variable (IntVar)
        gra_tax = gVar.get()
        pen_tax = pVar.get()
        ns_tax = n_sVar.get()
        nh_tax = n_hVar.get()

        if gra_tax == 1:
            gv = float(g_field.get())
            if not gv:
                gv = 0
        else:
            gv = 0

        if pen_tax == 1:
            pv = float(p_field.get())
            if not pv:
                pv = 0
        else:
            pv = 0

        if ns_tax == 1:
            nsv = float(ns_field.get())
            if not nsv:
                nsv = 0
        else:
            nsv = 0

        if nh_tax == 1:
            nhv = float(nh_field.get())
            if not nhv:
                nhv = 0
        else:
            nhv = 0

        # Variables to get info (Monthly gross income) from tree and calculate PAYE
        annVar = monVar * 12
        annVar2 = "{:,.2f}".format(annVar)  # To make the number have commas separating them depending on the value, and to have 2 decimal places. Note: this turns value to string

        cra = annVar * 0.2  # 20% of Annual Gross
        cra_Var = cra + 200000
        cra_Var2 = "{:,.2f}".format(cra_Var)  # To make the number have commas separating them depending on the value, and to have 2 decimal places. Note: this turns value to string

        taxVar = annVar - cra_Var - gv - pv - nsv - nhv
        taxVar2 = "{:,.2f}".format(taxVar)  # To make the number have commas separating them depending on the value, and to have 2 decimal places. Note: this turns value to string

        t = 43472  # Value to which 1 % of ann = 7% of tax income

        # if statements to sort out the annual tax pay
        if taxVar <= t:
            ann_tax_Var = annVar * 0.01
        elif t < taxVar <= 300000:
            ann_tax_Var = taxVar * 0.07
        elif 300000 < taxVar <= 600000:
            a = taxVar - 300000
            ann_tax_Var = ( a * 0.11 ) + 21000
        elif 600000 < taxVar <= 1100000:
            b = taxVar - 600000
            ann_tax_Var = ( b * 0.15 ) + 21000 + 33000
        elif 1100000 < taxVar <= 1600000:
            c = taxVar - 1100000
            ann_tax_Var = ( c * 0.19 ) + 21000 + 33000 + 75000
        elif 1600000 < taxVar <= 3200000:
            d = taxVar - 1600000
            ann_tax_Var = ( d * 0.21 ) + 21000 + 33000 + 75000 + 95000
        elif taxVar > 3200000:
            e = taxVar - 3200000
            ann_tax_Var = ( e * 0.24 ) + 21000 + 33000 + 75000 + 95000 + 336000
        else:
            pass


        # Mooore variables to solve our data
        mon_tax_Var = ann_tax_Var / 12
        mon_tax_Var2 = "{:,.2f}".format(mon_tax_Var)  # To make the number after commas separating depending on the value, and to have 2 decimal places. Note: this turns value to string

        ann_tax_Var2 = "{:,.2f}".format(ann_tax_Var)  # To make the number after commas separating depending on the value, and to have 2 decimal places. Note: this turns value to string



        g_varr2 = "{:,.2f}".format(gv)
        p_varr2 = "{:,.2f}".format(pv)
        ns_varr2 = "{:,.2f}".format(nsv)
        nh_varr2 = "{:,.2f}".format(nhv)

        # Insert values from tree (previous Treeview)
        tree.insert(parent="",
                    index="end",
                    id=count,
                    text="",
                    values=(serial, nameVar, desVar, tinVar, workVar, monVar2, annVar2, cra_Var2, p_varr2, nh_varr2, ns_varr2, g_varr2, taxVar2, ann_tax_Var2, mon_tax_Var2)
                    )

        count += 1  # Increment iid after every entry
        serial += 1  # Increment "S/N" after every entry

        main_frame.place_forget()
        g_frame.pack_forget()
        p_frame.pack_forget()
        ns_frame.pack_forget()
        nh_frame.pack_forget()

        g_field.delete(0, END)
        p_field.delete(0, END)
        ns_field.delete(0, END)
        nh_field.delete(0, END)

        g.deselect()
        p.deselect()
        n_s.deselect()
        n_h.deselect()

        exemp_add_btn['state'] = DISABLED






    else:

        # Variables to get info (Monthly gross income) from tree and calculate PAYE
        annVar = monVar * 12
        annVar2 = "{:,.2f}".format(annVar)  # To make the number after commas separating depending on the value, and to have 2 decimal places. Note: this turns value to string

        cra = annVar * 0.2  # 20% of Annual Gross
        cra_Var = cra + 200000
        cra_Var2 = "{:,.2f}".format(cra_Var)  # To make the number after commas separating depending on the value, and to have 2 decimal places. Note: this turns value to string

        taxVar = annVar - cra_Var
        taxVar2 = "{:,.2f}".format(taxVar)  # To make the number after commas separating depending on the value, and to have 2 decimal places. Note: this turns value to string

        t = 43472  # Value to which 1 % of ann = 7% of tax income

        # if statements to sort out the annual tax pay
        if taxVar <= t:
            ann_tax_Var = annVar * 0.01
        elif t < taxVar <= 300000:
            ann_tax_Var = taxVar * 0.07
        elif 300000 < taxVar <= 600000:
            a = taxVar - 300000
            ann_tax_Var = ( a * 0.11 ) + 21000
        elif 600000 < taxVar <= 1100000:
            b = taxVar - 600000
            ann_tax_Var = ( b * 0.15 ) + 21000 + 33000
        elif 1100000 < taxVar <= 1600000:
            c = taxVar - 1100000
            ann_tax_Var = ( c * 0.19 ) + 21000 + 33000 + 75000
        elif 1600000 < taxVar <= 3200000:
            d = taxVar - 1600000
            ann_tax_Var = ( d * 0.21 ) + 21000 + 33000 + 75000 + 95000
        elif taxVar > 3200000:
            e = taxVar - 3200000
            ann_tax_Var = ( e * 0.24 ) + 21000 + 33000 + 75000 + 95000 + 336000
        else:
            pass


        # Mooore variables to solve our data
        mon_tax_Var = ann_tax_Var / 12
        mon_tax_Var2 = "{:,.2f}".format(mon_tax_Var)  # To make the number after commas separating depending on the value, and to have 2 decimal places. Note: this turns value to string

        ann_tax_Var2 = "{:,.2f}".format(ann_tax_Var)  # To make the number after commas separating depending on the value, and to have 2 decimal places. Note: this turns value to string


        g_varr2 = "{:,.2f}".format(0)
        p_varr2 = "{:,.2f}".format(0)
        ns_varr2 = "{:,.2f}".format(0)
        nh_varr2 = "{:,.2f}".format(0)

        # Insert values from tree (previous Treeview)
        tree.insert(parent="",
                        index="end",
                        id=count,
                        text="",
                        values=(serial, nameVar, desVar, tinVar, workVar, monVar2, annVar2, cra_Var2, p_varr2, nh_varr2, ns_varr2, g_varr2, taxVar2, ann_tax_Var2, mon_tax_Var2)
                        )

        count += 1  # Increment iid after every entry
        serial += 1  # Increment "S/N" after every entry





    # Capitalize every inputted entry in "Organization entry box" (org_box) and store in variable
    capital_input = org_box.get().upper()

    # Replace "tree_frame"'s label with the capitalized entry
    sub_frame_label["text"] = capital_input + "'S" + " PAYE COMPUTATION"


    # To clear data from form after clicking button
    name_box.delete(0, END)
    des_box.delete(0, END)
    tin_box.delete(0, END)
    work_box.delete(0, END)
    mon_box.delete(0, END)

    # To deselect Checkbuttons
    g.deselect()
    p.deselect()
    n_s.deselect()
    n_h.deselect()


    # Disable after calling this function
    add_btn.configure(state=DISABLED)
    #edit_btn.configure(state=NORMAL)
    clear_btn.configure(state=NORMAL)
    org_box.configure(state=DISABLED)
    export_btn["state"]=NORMAL


# Function to upadte new entry in Treeview
def update():
    # variable to store the Serial "S/N" of the selected entry
    new_serial = values[0]
    # Grab entry data
    selected = tree.focus()
    # Save new data
    tree.item(selected, text="", values=(new_serial, new_name_box.get().upper(), new_des_box.get().upper(), new_tin_box.get(), new_work_box.get(), monVar2, annVar2, cra_Var2, p_varr2, nh_varr2, ns_varr2, g_varr2, taxVar2, ann_tax_Var2, mon_tax_Var2))

    # To close window making changes
    edit_wind.destroy()


# Function to select an Entry  and view it in a smaller pop-up window
def edit():
    # Make the new window a global variable to be able to use elsewhere
    global edit_wind
    # New window to display form for editing
    edit_wind = Tk()
    edit_wind.title("Edit Entry")
    edit_wind.geometry("480x300")
    edit_wind["bg"]="#d3d3d3"
    edit_wind.resizable(0, 0)

    form_frame = LabelFrame(edit_wind, text="")
    form_frame.place(height=250, width=430, relx=0.05, rely=0.065)
    form_frame["bg"]="#d3d3d3"


    # Label widgets for form
    new_name = Label(form_frame, text="Name", bg="#d3d3d3").place(relx=0.038, rely=0.1)
    new_des = Label(form_frame, text="Designation", bg="#d3d3d3").place(relx=0.038, rely=0.25)
    new_tin = Label(form_frame, text="T.I.N", bg="#d3d3d3").place(relx=0.038, rely=0.4)
    new_work = Label(form_frame, text="Months Worked", bg="#d3d3d3").place(relx=0.038, rely=0.55)
    #new_mon = Label(form_frame, text="Monthly Gross Income", bg="#d3d3d3").place(relx=0.038, rely=0.6)

    global new_name_box
    global new_des_box
    global new_tin_box
    global new_work_box
    global new_mon_box

    # Entry widgets for the form
    new_name_box = Entry(form_frame, width=25)
    new_name_box.place(relx=0.4, rely=0.1)

    new_des_box = Entry(form_frame, width=25)
    new_des_box.place(relx=0.4, rely=0.25)

    new_tin_box = Entry(form_frame, width=25)
    new_tin_box.place(relx=0.4, rely=0.4)

    new_work_box = Entry(form_frame, width=25)
    new_work_box.place(relx=0.4, rely=0.55)

    #new_mon_box = Entry(form_frame, width=25)
    #new_mon_box.place(relx=0.4, rely=0.6)


    # Buttons for the form
    update_btn = Button(form_frame, text="Update", width=10, command=update)
    update_btn.place(relx=0.2, rely=0.8)

    Button(form_frame, text="Cancel", width=10, command=edit_wind.destroy).place(relx=0.6, rely=0.8)

    global values

    # Grab no. of  selected Entry in Treeview
    selected = tree.focus()

    # Grab values of selected entry
    values = tree.item(selected, 'values')

    # Output the values of selected entry to the Entry widgets of "new_wind"
    new_name_box.insert(0, values[1])
    new_des_box.insert(0, values[2])
    new_tin_box.insert(0, values[3])
    new_work_box.insert(0, values[4])
    #new_mon_box.insert(0, values[5])

# Function to delete all entries in Treeview
def clear():
    global count
    global serial

    for record in tree.get_children():  # Function to get all components of treeview
        tree.delete(record)

    count = 0  # So that the treeview's id can reset
    serial = 1 # So that the serial number in treeciew, will start from 1 again

    # Change state of widgets when button is clicked
    edit_btn.configure(state=DISABLED)
    clear_btn.configure(state=DISABLED)
    org_box.configure(state=NORMAL)
    export_btn["state"]=DISABLED

    org_box.delete(0, END)  # Clear this entry widget when button is clicked
    sub_frame_label["text"]= "PAYE COMPUTATION"



    main_frame.place_forget()

    g_frame.pack_forget()
    p_frame.pack_forget()
    ns_frame.pack_forget()
    nh_frame.pack_forget()

    g_field.delete(0, END)
    p_field.delete(0, END)
    ns_field.delete(0, END)
    nh_field.delete(0, END)

    g.deselect()
    p.deselect()
    n_s.deselect()
    n_h.deselect()

    exemp_add_btn["state"] = DISABLED

# Function to export data from Treeview to Excel (.xlsx)
def export():

    #s = tree.get_children()

    #fln = filedialog.asksaveasfilename(initialdir=os.getcwd(), title="Save Computation", filetypes=(("Excel File", ".xlsx"),("All Files", "*.*")))

    #s.to_excel(fln)

    tree_columns = ["S/N", "NAME", "DESIGNATION", "TIN", "MONTHS WORKED", "MONTHLY GROSS INCOME", "ANNUAL GROSS INCOME", "RELIEF ALLOWANCE", "PENSION", "NATIONAL HOUSING FUND", "NHIS", "GRATUITIES", "TAXABLE INCOME", "ANNUAL TAX DUE", "MONTHLY TAX DUE"]
    treeview_df = pd.DataFrame(columns=tree_columns)

    for row in tree.get_children():
        values = pd.DataFrame([tree.item(row)["values"]], columns=tree_columns)

        treeview_df = treeview_df.append(values)


    #xlwt_write = pd.io.excel.get_writer("xlwt")
    engine = "xlsxwriter"
    save_file = filedialog.asksaveasfilename(defaultextension=".xlsx", initialdir=os.getcwd(), title="Save Computation", filetypes=(("Excel File", ".xlsx"),("All Files", "*.*")))
    writer = pd.ExcelWriter(save_file.format(engine), engine=engine)
    #writer = xlwt_writer(filedialog.asksaveasfilename(initialdir=os.getcwd(), title="Save Computation", filetypes=(("Excel File", ".xlsx"),("All Files", "*.*"))))
    treeview_df.to_excel(writer, index=False)
    writer.close()


# Button widget for the form
add_btn = Button(info_frame, text="Add Taxpayer", width=15, state=DISABLED)
add_btn.place(relx=0.15, rely=0.85)

export_btn = Button(info_frame, text="Export To MS Excel", width=15, state=DISABLED, command=export)
export_btn.place(relx=0.5, rely=0.85)


# Update and Delete Buttons for Treeview
edit_btn = Button(sub_frame, text="Edit Selected Entry", width=15, command=edit, state=DISABLED)
edit_btn.place(relx=0.275, rely=0.9)

clear_btn = Button(sub_frame, text="Clear Entry", width=15, command=clear, state=DISABLED)
clear_btn.place(relx=0.575, rely=0.9)

# Function binded to add_btn to change state to Normal
def clicker(event):
    add_btn.configure(state=NORMAL)

def edit_event(event):
    edit_btn.configure(state=NORMAL)

def check_box(event):
    exemp_add_btn.configure(state=NORMAL)




# Events binded to the last entry field of form and the "Add taxpayer" button
g.bind("<Button-1>", check_box) # To enable exemp_add_btn after clicking checkbox
p.bind("<Button-1>", check_box) # To enable exemp_add_btn after clicking checkbox
n_s.bind("<Button-1>", check_box) # To enable exemp_add_btn after clicking checkbox
n_h.bind("<Button-1>", check_box) # To enable exemp_add_btn after clicking checkbox
mon_box.bind("<FocusIn>", clicker)  # To enable add_btn when focused in through TAB button or mouse click
add_btn.bind("<Button-1>", add_payer) # To call add_payer() when use the Return button on keyboard
add_btn.bind("<Return>", add_payer) # To call add_payer()  after clicking the button
tree.bind("<Button-1>", edit_event) # To enable edit_btn after clicking Treeview


root.mainloop()
