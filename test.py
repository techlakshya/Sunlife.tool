from tkinter import *
from tkinter import ttk
from tkcalendar import *
import pyperclip
from tkinter import messagebox
class Coi_Calci:
    
    def __init__(lak,root):
        lak.root= root
        lak.root.title("COI Calulator")
        lak.root.geometry("460x500")
        lak.root.resizable(False,False)
        lak.var_province= StringVar()
        lak.var_plan= StringVar()
        lak.var_plan2= StringVar()
        lak.var_amount= StringVar()
        lak.Var_address= StringVar()
        lak.M_COI= StringVar()
        lak.A_COI= StringVar()
        lak.pla_search= StringVar()
        lak.Main_fram()

#--------------------------------------------------------------First Screen--------------------------------------------------------------------------------------------------    
    
    def Main_fram(lak):
        #main Fram
        lak.frame= Frame(lak.root)
        lak.frame.place (x=0,y=0,width=470,height=650)
        entry = Entry(lak.frame, relief=SUNKEN,borderwidth=10,foreground= "black",bg= "#fce8b0",font = ("Arial", 20, "bold"), width=28)
        entry.place(x=5,y=0)
        # Entrybox for plan combo box
        COI_entry =Entry(lak.frame, width=22,borderwidth=10,foreground= "black",font = ("Arial",12, "bold"))
        COI_entry.place(x=150,y=150)
        # Plan combo box
        Que_combo= ttk.Combobox(lak.frame, textvariable=lak.var_plan, width=20,font = ("Arial", 12, "bold"),state= "readonly")
        Que_combo["values"]= (" ","UL 2000", "SunSpectrum UL", "ULife Met","Flexible Premium Life", "VIAPP", "UL50","SunUL", "SunUL Max","SunUL II", 
        "SunUL Pro","Sun Ltd","SunSpectrum UL II")
        Que_combo.set(" ")
        Que_combo.place(x=160,y=157)
        #Label for plan combo box
        plan_lbl= Label(lak.frame, text= "Plan Type", font = ("Arial",12, "bold"), fg= "black")
        plan_lbl.place(x=10,y=153)
        # Entrybox for province combo box
        COI_entry =Entry(lak.frame, width=22,borderwidth=10,foreground= "black",font = ("Arial",12, "bold"))
        COI_entry.place(x=150,y=200)
        # province combo box
        Que_combo= ttk.Combobox(lak.frame, textvariable=lak.var_province, width=20,font = ("Arial",12, "bold"), state= "readonly")
        Que_combo["values"]=(" ","Non Residence","British Columbia" , "Manitoba" , "New Brunswick", "Labrador", "Ontario", "Yukon",
        "Alberta","Northwest Territories", "Nova Scotia" , "Nunavut","Quebec","Newfoundland","Prince Edward Island",
        "Saskatchewan(issued before 31 March,2000)","Saskatchewan(issued after 31 March,2000)")
        Que_combo.set(" ")
        Que_combo.place(x=160,y=210)
        #Label for province combo box
        Cprovince_lbl= Label(lak.frame, text= "Provinces", font = ("Arial",12, "bold"), fg= "black")
        Cprovince_lbl.place(x=10,y=203)
        #Label for amount
        province_lbl= Label(lak.frame, text= "Amount", font = ("Arial", 12, "bold"), fg= "black")
        province_lbl.place(x=10,y=253)
        # Entrybox for amount
        Eamount =Entry(lak.frame,textvariable=lak.var_amount, width=22,foreground= "black",font = ("Arial",12 , "bold"),borderwidth=10)
        Eamount.place(x=150,y=250)
        # Label for Monthly COI
        COI_lbl= Label(lak.frame, text= "Monthly COI", font = ("Arial", 12, "bold"), fg= "black")
        COI_lbl.place(x=10,y=305)
        # Entrybox for Monthly COI
        COI_entry =Entry(lak.frame,textvariable=lak.M_COI, width=22,foreground= "black",font = ("Arial",12, "bold"),borderwidth=10)
        COI_entry.place(x=150,y=300)
        # Label for Annually COI
        Annually_lbl= Label(lak.frame, text= "Annually COI", font = ("Arial",12, "bold"), fg= "black")
        Annually_lbl.place(x=10,y=355)
        # Entrybox for Annually COI
        Annually_entry =Entry(lak.frame,textvariable=lak.A_COI,  width=22,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        Annually_entry.place(x=150,y=350)

        # button
        button_method= Button(lak.frame, command=lak.method, activebackground= "#031b28",activeforeground= "white", 
        text = " Step ",fg= "black",width=6,font = ("Arial", 12, "bold"))
        button_method.place(x=380,y=150)
        button= Button(lak.frame, command=lak.Reset, activebackground= "#031b28",activeforeground= "white", 
        text = "  Reset   ",fg= "black",width=6,font = ("Arial", 12, "bold"))
        button.place(x=380,y=200)
        button= Button(lak.frame, command=lak.Coi_Calculator, activebackground= "#031b28",activeforeground= "white", 
        text = " Enter ",fg= "black",width=6,font = ("Arial", 12, "bold"))
        button.place(x=380,y=250)
        button4= Button(lak.frame, command=lambda:pyperclip.copy(lak.M_COI.get()), activebackground= "yellow",activeforeground= "white", 
        text = " Copy ",fg= "black",width=6,font = ("Arial", 12, "bold"))
        button4.place(x=380,y=300)
        button5= Button(lak.frame, command=lambda:pyperclip.copy(lak.A_COI.get()), activebackground= "yellow",activeforeground= "white", 
        text = " Copy ",fg= "black",width=6,font = ("Arial", 12, "bold"))
        button5.place(x=380,y=350)

        def myclick(number):
            entry.insert(END, number)
        def equal():
            try:
                y = str(eval(entry.get()))
                entry.delete(0, END)
                entry.insert(0, y)
            except:
                messagebox.showinfo("Error", "Enter the Number First")
        def clear():
            entry.delete(0, END)
        #button for calculator
        button_clear = Button(master=lak.frame, text="Clear", command=clear,width=18,activebackground= "#031b28", activeforeground= "white",
        fg= "black",font = ("Arial", 14, "bold"))
        button_clear.place(x=5,y=55)
        button_equal = Button(master=lak.frame, text="Equal",command=equal,width=18,activebackground= "#031b28", activeforeground= "white",
        fg= "black",font = ("Arial", 14, "bold"))
        button_equal.place(x=230,y=55)

#--------------------------------------------------------------Second Screen--------------------------------------------------------------------------------------
    def method(lak):
        if lak.var_plan.get() in ["SunUL", "SunUL Max","SunUL II", "SunUL Pro"]:
            step= """Amount=(Guaranted Cost of Insurance + Supplementry Benefit Charges + 
            Substandard Mortality Charges + RTB Charges)                                """
        else:
            step= "Amount= Total Deduction {:500} \n {:500}".format(" ", " ")
        # Label for Method
        Method_lbl= Label(lak.root, text=step, font = ("Arial", 10, "bold"), fg= "black")
        Method_lbl.place(x=0,y=430)

    def Reset(lak):
        lak.var_province.set("")
        lak.var_plan.set("")
        lak.var_amount.set("")
        lak.pla_search.set("")
        lak.M_COI.set("")
        lak.A_COI.set("")
        # Label 
        Method_lbl= Label(lak.root, text=" {:500} \n {:500}".format(" ", " "),width=400, borderwidth=0, font = ("Arial", 20, "bold"), fg= "black")
        Method_lbl.place(x=0,y=430)     
        return

    def Coi_Calculator(lak):
        province1 = ["British Columbia" , "Manitoba" , "New Brunswick", "Labrador", "Ontario", "Yukon","Saskatchewan(issued before 31 March,2000)"]
        province2 = ["Quebec"]
        province3= ["Newfoundland"]
        province4= ["Prince Edward Island"]
        province5= ["Alberta","Northwest Territories", "Nova Scotia" , "Nunavut", "Saskatchewan(issued after 31 March,2000)"]
        province6= ["Non Residence"]
        plan1= ["UL 2000", "SunSpectrum UL", "ULife Met"]
        plan2= ["Flexible Premium Life", "VIAPP", "UL50"]
        plan3= ["SunUL", "SunUL Max","SunUL II", "SunUL Pro"]
        plan4= ["Sun Ltd","SunSpectrum UL II"]
        p=lak.var_plan.get()
        s=lak.var_province.get()
        x=float(lak.var_amount.get())

        if p in plan1:
            if s in province1:
                Month_COI= x/0.98
                Annual_COI= (Month_COI * 12)            
            elif s in province2:
                Month_COI= (x/0.967)
                Annual_COI= (Month_COI * 12)            
            elif s in province3:
                Month_COI= (x/0.96)
                Annual_COI= (Month_COI * 12)            
            elif s in province4:
                Month_COI= (x/0.965)
                Annual_COI= (Month_COI * 12)            
            elif s in province5:
                Month_COI= x/0.97
                Annual_COI= (Month_COI * 12)            
            elif s in province6:
                Month_COI= x/0.98
                Annual_COI= (Month_COI * 12)            
            else:
                Month_COI= x/0.98
                Annual_COI= (Month_COI * 12)            
        elif p in plan2:
            Month_COI= (x * 1.15)
            Annual_COI= (Month_COI * 12)            
        elif p in plan3:
            if s in province1:
                Month_COI= x/0.98
                Annual_COI= (Month_COI * 12)            
            elif s in province2:
                Month_COI= (x/0.967)
                Annual_COI= (Month_COI * 12)            
            elif s in province3:          
                Month_COI= (x/0.96)
                Annual_COI= (Month_COI * 12)            
            elif s in province4:
                Month_COI= (x/0.965)
                Annual_COI= (Month_COI * 12)            
            elif s in province5:
                Month_COI= x/0.97
                Annual_COI= (Month_COI * 12)            
            elif s in province6:
                x=lak.var_amount.get()
                Month_COI= x/0.98
                Annual_COI= (Month_COI * 12)            
            else:
                Month_COI= x/0.98
                Annual_COI= (Month_COI * 12)            
        elif p in plan4:
            Month_COI= x/0.98
            Annual_COI= (Month_COI * 12)            
        lak.M_COI.set(round(Month_COI, 2))
        lak.A_COI.set(round(Annual_COI, 2))


if __name__== "__main__":
    root=Tk()
    aap= Coi_Calci(root)
    root.mainloop()