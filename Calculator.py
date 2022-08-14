from telnetlib import DO
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from PIL import Image,ImageTk

class Coi_Calci:
    
    def __init__(lak,root):
        lak.root= root
        lak.root.title("COI Calulator")
        lak.root.geometry("600x800")
        lak.root.resizable(False,False)
        lak.root.wm_iconbitmap("sunlife.ico")
        lak.var_province= StringVar()
        lak.var_plan= StringVar()
        lak.var_amount= StringVar()
        # backgroung image
        lak.bg=ImageTk.PhotoImage(file="E:\\Sunlife\\slbg.jpg")
        label_bg=Label(lak.root, image=lak.bg,borderwidth=0)
        label_bg.place(x=0,y=0,relwidth=1,relheight=1)
#--------------------------------------------------------------Frams--------------------------------------------------------------------------------------
        # frame
        lak.frame= Frame(lak.root,  bg= "white")
        lak.frame.place (x=20,y=400,width=600,height=300)
        # button frame
        lak.frame2= Frame(lak.root,  bg= "white")
        lak.frame2.place (x=50,y=660,width=550,height=50)
        # method frame
        lak.frame3= Frame(lak.root, bg= "white")
        lak.frame3.place (x=20,y=200,width=310,height=200)
        #Label for Method
        plan_lbl= Label(lak.frame3,text= '''
        Method A: 
        Amount will be the sum of
        (Guaranted Cost of Insurance,
        Supplementry Benefit Charges
        Substandard Mortality Charges,
        RTB Charges)
        
        Method B:
        Amount will be the 
        Total Deduction amount:''',
        bg= "white", font = ("Arial", 8, "bold"), fg= "black")
        plan_lbl.place(x=0,y=10)
#--------------------------------------------------------------Combobox--------------------------------------------------------------------------------------
        # Entrybox for province combo box
        COI_entry =Entry(lak.frame, width=30,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        COI_entry.place(x=100,y=50)
        # province combo box
        Que_combo= ttk.Combobox(lak.frame, textvariable=lak.var_province, width=31,font = ("Arial", 11, "bold"), state="readonly")
        Que_combo["values"]= (" ","Non Residence","British Columbia" , "Manitoba" , "New Brunswick", "Labrador", "Ontario", "Yukon",
        "Alberta","Northwest Territories", "Nova Scotia" , "Nunavut","Quebec","Newfoundland","Prince Edward Island","Saskatchewan(issued before 31 March,2000)","Saskatchewan(issued after 31 March,2000)")
        Que_combo.set(" ")
        Que_combo.place(x=110,y=60)
        #Label for province combo box
        Cprovince_lbl= Label(lak.frame, text= "Provinces:",bg= "white", font = ("Arial", 12, "bold"), fg= "black")
        Cprovince_lbl.place(x=10,y=50)
        # Entrybox for plan combo box
        COI_entry =Entry(lak.frame, width=30,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        COI_entry.place(x=100,y=0)
        # Plan combo box
        Que_combo= ttk.Combobox(lak.frame, textvariable=lak.var_plan, width=31,font = ("Arial", 11, "bold"), state="readonly")
        Que_combo["values"]= (" ","UL 2000", "SunSpectrum UL", "ULife Met","Flexible Premium Life", "VIAPP", "UL50","SunUL", "SunUL Max","SunUL II", 
        "SunUL Pro","Sun Ltd","SunSpectrum UL II")
        Que_combo.set(" ")
        Que_combo.place(x=110,y=10)
        #Label for plan combo box
        plan_lbl= Label(lak.frame, text= "Plan Type:",bg= "white", font = ("Arial", 12, "bold"), fg= "black")
        plan_lbl.place(x=10,y=0)
#--------------------------------------------------------------Entry/Lable--------------------------------------------------------------------------------------
        #Label for amount
        province_lbl= Label(lak.frame, text= "Amount   :",bg= "white", font = ("Arial", 12, "bold"), fg= "black")
        province_lbl.place(x=10,y=100)
        # Entrybox for amount
        Eamount =Entry(lak.frame,textvariable=lak.var_amount, width=25,foreground= "black",font = ("Arial", 14, "bold"),borderwidth=10)
        Eamount.place(x=100,y=100)
        # Label for Monthly COI
        COI_lbl= Label(lak.frame, text= "Monthly COI :",width= 10, borderwidth=10, font = ("Arial", 12, "bold"), fg= "#008000", bg= "white")
        COI_lbl.place(x=150,y=150)
        # Entrybox for Monthly COI
        COI_entry =Entry(lak.frame, width=20,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        COI_entry.place(x=300,y=150)
        # Label for Annually COI
        Annually_lbl= Label(lak.frame, text= "Annually COI :",width= 10, borderwidth=10, font = ("Arial", 12, "bold"), fg= "#008000", bg= "white")
        Annually_lbl.place(x=150,y=200)
        # Entrybox for Annually COI
        Annually_entry =Entry(lak.frame, width=20,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        Annually_entry.place(x=300,y=200)
#--------------------------------------------------------------Bottons--------------------------------------------------------------------------------------
        # Coi button
        button= Button(lak.frame2, command=lak.Coi_Calculator, activebackground= "#031b28",activeforeground= "white", 
        text = " COI Calculation ",fg= "white",width=15, bg= "#031b28",font = ("Arial", 12, "bold italic"))
        button.place(x=170,y=10)
        # reset button
        button= Button(lak.frame2, command=lak.Reset, activebackground= "#031b28",activeforeground= "white", 
        text = "  Reset   ",fg= "white", bg= "#031b28",width=15,font = ("Arial", 12, "bold italic"))
        button.place(x=340,y=10)
        # Method button
        Mbutton= Button(lak.frame2, command= lak.Method, activebackground= "#031b28",width=15, activeforeground= "white", 
        text = "   Method   ",fg= "white", bg= "#031b28",font = ("Arial", 12, "bold italic"))
        Mbutton.place(x=0,y=10)
#--------------------------------------------------------------Reset()--------------------------------------------------------------------------------------
    def Reset(lak):
        lak.var_province.set("")
        lak.var_plan.set("")
        lak.var_amount.set("")
        return
#--------------------------------------------------------------Method()--------------------------------------------------------------------------------------        
    def Method(lak):
        Me_plan3= ["SunUL", "SunUL Max","SunUL II", "SunUL Pro"]
        p=lak.var_plan.get()
        if p in Me_plan3:
            messagebox.showinfo("Method", "For Amount PLease Follow Method A")
        else:
            messagebox.showinfo("Method", "For Amount PLease Follow Method B")
#--------------------------------------------------------------main()--------------------------------------------------------------------------------------            
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
        Monthly_lbl= Label(lak.frame, text= Month_COI ,width= 20, borderwidth=0, font = ("Arial", 11, "bold"), fg= "black", bg= "white")
        Monthly_lbl.place(x=310,y=158)
        Annually_lbl= Label(lak.frame, text= Annual_COI ,width= 20, borderwidth=0, font = ("Arial", 11, "bold"), fg= "black", bg= "white")
        Annually_lbl.place(x=310,y=208)
#--------------------------------------------------------------Object()--------------------------------------------------------------------------------------
if __name__== "__main__":
    root=Tk()
    aap= Coi_Calci(root)
    root.mainloop()