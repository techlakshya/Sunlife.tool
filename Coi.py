from tkinter import *
from tkinter import ttk
from tkcalendar import *
from tkcalendar import DateEntry
import pyperclip
import pandas as pd
from pandas import DataFrame as df
import holidays
from datetime import datetime
from tkinter import messagebox

class Coi_tool:
    
    def __init__(lak,root):
        lak.root= root
        lak.root.title("Amount Canculation Tool")
        lak.root.geometry("800x950")
        lak.root.resizable(False,False)
        lak.var_province= StringVar()
        lak.var_plan= StringVar()
        lak.var_plan2= StringVar()
        lak.var_amount= StringVar()
        lak.Var_address= StringVar()
        lak.address1= StringVar()
        lak.address2= StringVar()
        lak.city= StringVar()
        lak.Province= StringVar()
        lak.Zip_code= StringVar()
        lak.pla_search= StringVar()
        lak.M_COI= StringVar()
        lak.A_COI= StringVar()
        lak.var_prems_amt= StringVar()
        lak.var_susp_amt= StringVar()
        lak.var_accum_amt= StringVar()
        lak.var_nsf_amt= StringVar()
        lak.var_outstand_amt= StringVar()
        lak.var_p_to_d= StringVar()
        lak.var_n_with_d= StringVar()
        lak.var_province2= StringVar()
        lak.bday_date= StringVar()
        lak.bday_date10= StringVar()
        lak.bday_date20= StringVar()
        lak.Lday_date62= StringVar()
        lak.number_days= StringVar()
        lak.w_freq= StringVar()
        lak.w_policy_val= StringVar()
        lak.w_Anvercry_date= StringVar()
        lak.w_COI= StringVar()
        lak.w_left= StringVar()
        lak.w_with_amot= StringVar()
        lak.w_hiddin_date= StringVar()
        lak.Main_fram()
#--------------------------------------------------------------First Screen--------------------------------------------------------------------------------------------------    
    
    def Main_fram(lak):
        #main Fram
        label_bg=Label(lak.root ,borderwidth=0, background="#53085C")
        label_bg.place(x=0,y=0,relwidth=1,relheight=1)
        # Nevigation button
        Mbutton2= Button(lak.root, command= lak.Main_fram, activebackground= "#031b28",width=35, activeforeground= "white", 
        text = "COI Amount Caculation",fg= "white", bg= "#53085C",font = ("Arial", 15, "bold "))
        Mbutton2.place(x=0,y=0)
        Mbutton2= Button(lak.root, command= lak.Withdrawal_Amount, activebackground= "#031b28",width=35, activeforeground= "white", 
        text = "Withdrawable/Calander",fg= "black", bg= "white",font = ("Arial", 15, "bold "))
        Mbutton2.place(x=400,y=0)
        #hidden
        COI_entry =Entry(lak.root, width=40,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        COI_entry.place(x=250,y=150)
        COI_entry =Entry(lak.root, width=40,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        COI_entry.place(x=250,y=200)
        COI_entry =Entry(lak.root, width=20,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        COI_entry.place(x=190,y=627)
        PTD_lbl_entry =Entry(lak.root,  width=20,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        PTD_lbl_entry.place(x=578,y=625)
        PTD_lbl_entry =Entry(lak.root, width=20,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        PTD_lbl_entry.place(x=190,y=680)
        NWD_entry =Entry(lak.root, width=20,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        NWD_entry.place(x=578,y=680)
        #Lable
        titl_lbl= Label(lak.root,relief="raise" ,width= 80,justify="left", text= "Cost Of Insurance Amount", font = ("Arial", 12, "bold"), fg= "white", bg= "#3E0A44")
        titl_lbl.place(x=0,y=47,height=40)
        prod_lbl= Label(lak.root,width= 20,justify="left", borderwidth=10, text= "Product Type",bg= "#3E0A44", font = ("Arial", 12, "bold"), fg= "white")
        prod_lbl.place(x=10,y=100)
        plan_lbl= Label(lak.root,width= 20,justify="left", borderwidth=10, text= " Plan Type",bg= "#3E0A44", font = ("Arial", 12, "bold"), fg= "white")
        plan_lbl.place(x=10,y=150)
        Cprovince_lbl= Label(lak.root,width= 20,justify="left", borderwidth=10, text= "Provinces",bg= "#3E0A44", font = ("Arial", 12, "bold"), fg= "white")
        Cprovince_lbl.place(x=10,y=200)
        province_lbl= Label(lak.root,width= 20,justify="left", borderwidth=10, text= "Amount",bg= "#3E0A44", font = ("Arial", 12, "bold"), fg= "white")
        province_lbl.place(x=10,y=250)
        COI_lbl= Label(lak.root,justify="left", text= "Monthly COI",width= 20, borderwidth=10, font = ("Arial", 12, "bold"), fg= "white", bg= "#3E0A44")
        COI_lbl.place(x=10,y=300)
        Annually_lbl= Label(lak.root,justify="left", text= "Annually COI",width= 20, borderwidth=10, font = ("Arial", 12, "bold"), fg= "white", bg= "#3E0A44")
        Annually_lbl.place(x=10,y=350)
        titl_lbl= Label(lak.root,relief="raise" ,width= 80,justify="left", text= "Outstanding Amount", font = ("Arial", 12, "bold"), fg= "white", bg= "#3E0A44")
        titl_lbl.place(x=0,y=480,height=40)
        #Entry
        P_entry =Entry(lak.root,textvariable= lak.pla_search, width=40,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        P_entry.place(x=250,y=100)
        Que_combo= ttk.Combobox(lak.root, textvariable=lak.var_plan, width=42,font = ("Arial", 11, "bold"),state= "readonly")
        Que_combo["values"]= (" ","UL 2000", "SunSpectrum UL", "ULife Met","Flexible Premium Life", "VIAPP", "UL50","SunUL", "SunUL Max","SunUL II", 
        "SunUL Pro","Sun Ltd","SunSpectrum UL II")
        Que_combo.set(" ")
        Que_combo.place(x=260,y=160)
        Que_combo= ttk.Combobox(lak.root, textvariable=lak.var_province, width=42,font = ("Arial", 11, "bold"), state= "readonly")
        Que_combo["values"]=(" ","Non Residence","British Columbia" , "Manitoba" , "New Brunswick", "Labrador", "Ontario", "Yukon",
        "Alberta","Northwest Territories", "Nova Scotia" , "Nunavut","Quebec","Newfoundland","Prince Edward Island",
        "Saskatchewan(issued before 31 March,2000)","Saskatchewan(issued after 31 March,2000)")
        Que_combo.set(" ")
        Que_combo.place(x=260,y=210)
        Eamount =Entry(lak.root,textvariable=lak.var_amount, width=40,foreground= "black",font = ("Arial", 12, "bold"),borderwidth=10)
        Eamount.place(x=250,y=250)
        COI_entry =Entry(lak.root,textvariable=lak.M_COI, width=40,foreground= "black",font = ("Arial", 12, "bold"),borderwidth=10)
        COI_entry.place(x=250,y=300)
        Annually_entry =Entry(lak.root,textvariable=lak.A_COI,  width=40,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        Annually_entry.place(x=250,y=350)
        # button
        button= Button(lak.root, command=lak.plan_search, activebackground= "#031b28",activeforeground= "white", 
        text = " Search ",width=12,fg= "white", bg= "#3E0A44",font = ("Arial", 12, "bold"),borderwidth=5)
        button.place(x=650,y=100)
        button_method= Button(lak.root, command=lak.method, activebackground= "#031b28",activeforeground= "white", 
        text = " Step ",width=12,fg= "white", bg= "#3E0A44",font = ("Arial", 12, "bold"),borderwidth=5)
        button_method.place(x=650,y=150)
        button= Button(lak.root, command=lak.Reset, activebackground= "#031b28",activeforeground= "white", 
        text = "  Reset ",width=12,fg= "white", bg= "#3E0A44",font = ("Arial", 12, "bold"),borderwidth=5)
        button.place(x=650,y=200)
        button= Button(lak.root, command=lak.Coi_Calculator, activebackground= "#031b28",activeforeground= "white", 
        text = " Enter ",width=12,fg= "white", bg= "#3E0A44",font = ("Arial", 12, "bold"),borderwidth=5)
        button.place(x=650,y=250)
        button4= Button(lak.root, command=lambda:pyperclip.copy(lak.M_COI.get()), activebackground= "yellow",activeforeground= "white", 
        text = " Copy ",width=12,fg= "white", bg= "#3E0A44",font = ("Arial", 12, "bold"),borderwidth=5)
        button4.place(x=650,y=300)
        button5= Button(lak.root, command=lambda:pyperclip.copy(lak.A_COI.get()), activebackground= "yellow",activeforeground= "white", 
        text = " Copy ",width=12,fg= "white", bg= "#3E0A44",font = ("Arial", 12, "bold"),borderwidth=5)
        button5.place(x=650,y=350)
        button= Button(lak.root, command=lak.OutStanding, activebackground= "#031b28",activeforeground= "white", 
        text = " Enter ",width=15,fg= "white", bg= "#3E0A44",font = ("Arial", 14, "bold"),borderwidth=5)
        button.place(x=585,y=837)
        #Calculator
        entry = Entry(lak.root,bg="#C681D0", relief=SUNKEN,borderwidth=10,foreground= "black",font = ("Arial", 24, "bold"), width=43)
        entry.place(x=0,y=520)
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
        #calculator
        button_clear = Button(master=lak.root, text="Clear", command=clear,width=28,activebackground= "#031b28", activeforeground= "white",
        fg= "white", bg= "#3E0A44",font = ("Arial", 12, "bold"))
        button_clear.place(x=0,y=578)
        button_equal = Button(master=lak.root, text="Equal",command=equal,width=28,activebackground= "#031b28", activeforeground= "white",
        fg= "white", bg= "#3E0A44",font = ("Arial", 12, "bold"))
        button_equal.place(x=265,y=578)
        button_equal = Button(master=lak.root, text="Rest",command=lak.Os_reset,width=28,activebackground= "#031b28", activeforeground= "white",
        fg= "white", bg= "#3E0A44",font = ("Arial", 12, "bold"))
        button_equal.place(x=520,y=578)

        #Outstanding
        Cprovince_lbl= Label(lak.root,justify="left", text= "Provinces",width= 15, borderwidth=10, font = ("Arial", 12, "bold"), fg= "white", bg= "#3E0A44")
        Cprovince_lbl.place(x=10,y=627)
        Que_combo= ttk.Combobox(lak.root, textvariable=lak.var_province2, width=20,font = ("Arial", 11, "bold"), state= "readonly")
        Que_combo["values"]=(" ","Non Residence","British Columbia" , "Manitoba" , "New Brunswick", "Labrador", "Ontario", "Yukon",
        "Alberta","Northwest Territories", "Nova Scotia" , "Nunavut","Quebec","Newfoundland","Prince Edward Island",
        "Saskatchewan(issued before 31 March,2000)","Saskatchewan(issued after 31 March,2000)")
        Que_combo.set(" ")
        Que_combo.place(x=197,y=635)
        plan_lbl1= Label(lak.root,justify="left", text= "Plan Type",width= 15, borderwidth=10, font = ("Arial", 12, "bold"), fg= "white", bg= "#3E0A44")
        plan_lbl1.place(x=400,y=625)
        Que_combo2= ttk.Combobox(lak.root, textvariable=lak.var_plan2, width=20,font = ("Arial", 11, "bold"),state= "readonly")
        Que_combo2["values"]= (" ","UL 2000", "SunSpectrum UL", "ULife Met","Flexible Premium Life", "VIAPP", "UL50","SunUL", "SunUL Max","SunUL II", 
        "SunUL Pro","Sun Ltd","SunSpectrum UL II")
        Que_combo2.set(" ")
        Que_combo2.place(x=585,y=633)
        PTD_lbl= Label(lak.root,justify="left", text= "Paid to date",width= 15, borderwidth=10, font = ("Arial", 12, "bold"), fg= "white", bg= "#3E0A44")
        PTD_lbl.place(x=10,y=680)
        cal= DateEntry(lak.root,selectmode="day",textvariable=lak.var_p_to_d, borderwidth=0,foreground= "black",font = ("Arial",12, "bold"), width=18)
        cal.place(x=200,y=685)
        NWD_lbl= Label(lak.root,justify="left", text= "Next Withdrawal",width= 15, borderwidth=10, font = ("Arial", 12, "bold"), fg= "white", bg= "#3E0A44")
        NWD_lbl.place(x=400,y=680)
        cal= DateEntry(lak.root,selectmode="day",textvariable=lak.var_n_with_d, borderwidth=0,foreground= "black",font = ("Arial",12, "bold"), width=18)
        cal.place(x=588,y=685)
        Premium_lbl= Label(lak.root,justify="left", text= "Min. Premium Amt.",width= 15, borderwidth=10, font = ("Arial", 12, "bold"), fg= "white", bg= "#3E0A44")
        Premium_lbl.place(x=10,y=735)
        Premium_entry =Entry(lak.root,textvariable=lak.var_prems_amt,  width=20,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        Premium_entry.place(x=190,y=735)
        Suspense= Label(lak.root,justify="left", text= "Suspense Amt.",width= 15, borderwidth=10, font = ("Arial", 12, "bold"), fg= "white", bg= "#3E0A44")
        Suspense.place(x=400,y=735)
        Suspense_lbl_entry =Entry(lak.root,textvariable=lak.var_susp_amt,  width=20,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        Suspense_lbl_entry.place(x=578,y=735)
        Accumulated_lbl= Label(lak.root,justify="left", text= "Accumulated Amt.",width= 15, borderwidth=10, font = ("Arial", 12, "bold"), fg= "white", bg= "#3E0A44")
        Accumulated_lbl.place(x=10,y=785)
        Accumulated_lbl_entry =Entry(lak.root,textvariable=lak.var_accum_amt,  width=20,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        Accumulated_lbl_entry.place(x=190,y=785)
        Fee_lbl= Label(lak.root,justify="left", text= "NSF Fee",width= 15, borderwidth=10, font = ("Arial", 12, "bold"), fg= "white", bg= "#3E0A44")
        Fee_lbl.place(x=400,y=785)
        Fee_lbl_entry =Entry(lak.root,textvariable=lak.var_nsf_amt,  width=20,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        Fee_lbl_entry.place(x=578,y=785)
        outs_lbl= Label(lak.root,justify="left", text= "Outstanding Amt",width= 15, borderwidth=10, font = ("Arial", 12, "bold"), fg= "white", bg= "#3E0A44")
        outs_lbl.place(x=10,y=840)
        outs_lbl_entry =Entry(lak.root,textvariable=lak.var_outstand_amt, bg="#C681D0", width=40,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        outs_lbl_entry.place(x=190,y=840)

    def plan_search(lak):
        SunLtd=['SL Sun Ltd SLP101', 'SL Sun Ltd SLP102', 'SL Sun Ltd SLP151', 'SL Sun Ltd SLP152', 'SL Sun Ltd SLP201', 'SL Sun Ltd SLP202', 'SL Sun Ltd SLP651', 
        'SL Sun Ltd JLP101', 'SL Sun Ltd JLP102', 'SL Sun Ltd JLP151', 'SL Sun Ltd JLP201', 'SL Sun Ltd JLP202',"Sun Ltd","SunSpectrum UL II"]
        UL50=['Interest Plus', 'ZIP', 'Flexible Premium Life', 'UL50', 'Universal Life 50', 'VIAPP', 'Versatile Investment and Protection Plan']
        ULife_Met=['Universal Plus','Univ Plus pre92 ', 'Univ Plus' 'Universal Plus joint', "Universal Plus (pre '92)", "Universal Plus joint (pre '92)Universal Flexiplus", 
        'Universal flexiplus joint', 'Universal Optimet', 'Universal Optimet joint', 'Interest Plus','SunSpectrum UL', 'SunSpectrum Universal Life', 'Jt. SunSpect UL', 
        'SunSpectrum Universal Life - Joint', 'SunSpectrum UL Increases', 
        'SunSpectrum Universal Life - Increase', 'Joint SunSpectrum UL Increases', 'SunSpectrum Universal Life - Joint Increase', 'Univ Life', 'Universal Life', 
        'UL - Splss', 'Universal Life - Special Issue', 'UL - SplssJt', 'Universal Life - Special Issue Joint', 'UL Inc (Universal Life - Increase)', 'UL Splss In', 
        'Universal Life - Special Issue Increase', 'UL Jt Inc', 'Universal Life - Joint Increase', 'Executive ULife',"UL 2000", "SunSpectrum UL", "ULife Met"]
        SunUL =["SunUL", "SunUL Max","SunUL II", "SunUL Pro",'SL ULA 1', 'SL ULA 2', 'SL ULA 3', 'SL ULA 4', 'SL ULA 5', 'SL ULA 6', 'SL ULA 7', 'SL ULA 8', 'SL ULA 9', 
        'SL ULJ 1', 'SL ULJ 2', 'SL ULJ 3','SL ULJ 4', 'SL ULJ 5', 'SL ULJ 6', 'SL ULJ 7', 'SL ULJ 8', 'SL ULJ 9', 'SL ULJ12 1', 'SL ULJ12 2', 'SL ULJ12 3', 'SL ULJ12 4', 
        'SL ULJ12 5', 'SL ULJ12 6','SL ULJ12 7', 'SL ULJ12 8', 'SL ULJ12 9', 'SL ULJ13 1', 'SL ULJ13 2', 'SL ULJ13 3', 'ULJ13 4', 'SL ULJ13 5', 'SL ULJ13 6', 'SL ULJ13 7', 
        'SL ULJ13 8','SL ULJ13 9', 'SL ULJ2 1', 'SL ULJ2 2', 'SL ULJ2 3', 'SL ULJ2 4', 'SL ULJ2 5', 'SL ULJ2 6', 'SL ULJ2 7', 'SL ULJ2 8', 'SL ULJ2 9', 'SL ULAX 1', 'SL ULAX 2', 
        'SL ULAX 3', 'SL ULAX 4', 'SL ULAX 5', 'SL ULAX 6', 'SL ULAX 7', 'SL ULAX 8', 'SL ULAX 9', 'SL SULX 1', 'SOLX 1', 'SL ULJX 1', 'ULJX 2', 'ULJX 3', 'ULJX 4', 
        'ULJX 5', 'ULJX 6', 'ULJX 7', 'ULJX 8', 'ULJX 9', 'SL UJ12X 1', 'UJ12X 2', 'UJ12X 3', 'UJ12X 4', 'UJ12X 5', 'UJ12X 6', 'UJ12X 7', 'UJ12X 8', 'UJ12X 9', 
        'SL UJ13X 1', 'UJ13X 2', 'UJ13X 3', 'UJ13X 4', 'UJ13X 5', 'UJ13X 6', 'UJ13X 7', 'UJ13X 8', 'UJ13X 9', 'SL UJ2X 1', 'UJ2X 2', 'UJ2X 3', 'UJ2X 4', 
        'UJ2X 5', 'UJ2X 6', 'UJ2X 7', 'UJ2X 8', 'UJ2X 9', 'SL ULC1 1', 'ULC1 2', 'ULC1 3', 'SL ULC1 1', 'ULC1 2', 'ULC1 3', 'SL ULC2 1', 'ULC2 2', 'ULC2 3', 
        'SL ULC3 1', 'ULC3 2', 'ULC3 3', 'SL SUL 1', 'SL SUJ 1', 'SL SUJ121', 'SL SUJ131', 'SL SUJ2 1', 'SL SEL 1', 'SL SEJ 1', 'SL SEJ121', 'SL SEJ131', 
        'SL SEJ2 1', 'SL SULX 1', 'SL SUJX 1', 'SL SU12X1', 'SL SU13X1', 'SL SU2X 1', 'SL SELX 1', 'SL SEJX 1', 'SL SE12X1', 'SL SE13X1', 'SL SE2X 1', 'SL AUL 1', 
        'SL AUJ 1','SL AUJ121', 'SL AUJ131', 'SL AUJ2 1', 'SL ULA A', 'ULA B', 'ULA C', 'ULA D', 'ULA E', 'ULA F', 'ULA G', 'ULA H', 'SL ULJ A', 'ULJ B', 'ULJ C', 
        'ULJ D', 'ULJ E', 'ULJ F', 'ULJ G', 'ULJ H', 'SL ULJ12 A', 'ULJ12 B', 'ULJ12 C', 'ULJ12 D', 'ULJ12 E', 'ULJ12 F', 'ULJ12 G', 'ULJ12 H', 'SL ULJ13 A', 
        'ULJ13 B', 'ULJ13 C', 'ULJ13 D', 'ULJ13 E', 'ULJ13 F', 'ULJ13 G', 'ULJ13 H', 'SL ULJ2 A', 'ULJ2 B', 'ULJ2 C', 'ULJ2 D', 'ULJ2 E', 'ULJ2 F', 'ULJ2 G', 
        'ULJ2 H','SL SOL 1','SL SOJ 1','SL SOJ121', 'SL SOJ131', 'SL SOJ2 1','SL EOL 1', 'SL EOJ 1','SL EOJ121','SL EOJ131','SL EOJ2 1','SL SOLX 1', 
        'SL SOJX 1', 'SL SO12X1', 'SL SO13X1', 'SL SO2X 1', 'SL EOLX 1', 'SL EOJX 1', 'SL EO12X1', 'SL EO13X1', 'SL EO2X 1', 'SL AOL 1', 'SL AOJ 1', 'SL AOJ121', 
        'SL AOJ131', 'SL AOJ2 1', 'SL ULC1 1', 'ULC1 2', 'ULC1 3', 'SL ULC2 1', 'ULC2 2', 'ULC2 3', 'SL ULC3 1', 'ULC3 2', 'ULC3 3',
        'SL UJ12X 2', 'SL UJ12X 3', 'SL UJ12X 4', 'SL UJ12X 5', 'SL UJ12X 6', 'SL UJ12X 7', 'SL UJ12X 8', 'SL UJ12X 9', 
        'SL UJ13X 1', 'SL UJ13X 2', 'SL UJ13X 3', 'SL UJ13X 4', 'SL UJ13X 5', 'SL UJ13X 6', 'SL UJ13X 7', 'SL UJ13X 8', 'SL UJ13X 9', 'SL UJ2X 1', 'SL UJ2X 2', 
        'SL UJ2X 3', 'SL UJ2X 4', 'SL UJ2X 5', 'SL UJ2X 6', 'SL UJ2X 7', 'SL UJ2X 8', 'SL UJ2X 9', 'SL ULC1 1', 'SL ULC1 2', 'SL ULC1 3', 'SL ULC1 1', 'SL ULC1 2', 
        'SL ULC1 3', 'SL ULC2 1', 'SL ULC2 2', 'SL ULC2 3', 'SL ULC3 1', 'SL ULC3 2', 'SL ULC3 3', 'SL SUL 1', 'SL SUJ 1', 'SL SUJ121', 'SL SUJ131', 'SL SUJ2 1', 'SL SEL 1', 
        'SL SEJ 1', 'SL SEJ121', 'SL SEJ131', 'SL SEJ2 1', 'SL SULX 1', 'SL SUJX 1', 'SL SU12X1', 'SL SU13X1', 'SL SU2X 1', 'SL SELX 1', 
        'SL SEJX 1', 'SL SE12X1', 'SL SE13X1', 'SL SE2X 1', 'SL AUL 1', 'SL AUJ 1', 'SL AUJ121', 'SL AUJ131', 
        'SL AUJ2 1', 'SL ULA A', 'SL ULA B', 'SL ULA C', 'SL ULA D', 'SL ULA E', 'SL ULA F', 'SL ULA G', 'SL ULA H', 'SL ULJ A', 'SL ULJ B', 
        'SL ULJ C', 'SL ULJ D', 'SL ULJ E', 'SL ULJ F', 'SL ULJ G', 'SL ULJ H', 'SL ULJ12 A', 'SL ULJ12 B', 'SL ULJ12 C', 'SL ULJ12 D', 'SL ULJ12 E', 
        'SL ULJ12 F', 'SL ULJ12 G', 'SL ULJ12 H', 'SL ULJ13 A', 'SL ULJ13 B', 'SL ULJ13 C', 'SL ULJ13 D', 'SL ULJ13 E', 'SL ULJ13 F', 'SL ULJ13 G', 'SL ULJ13 H', 
        'SL ULJ2 A', 'SL ULJ2 B', 'SL ULJ2 C', 'SL ULJ2 D', 'SL ULJ2 E', 'SL ULJ2 F', 'SL ULJ2 G', 'SL ULJ2 H', 'SL SOL 1', 'SL SOJ 1', 'SL SOJ121', 'SL SOJ131', 
        'SL SOJ2 1', 'SL EOL 1', 'SL EOJ 1', 'SL EOJ121', 'SL EOJ131', 'SL EOJ2 1', 'SL SOLX 1', 'SL SOJX 1', 'SL SO12X1', 'SL SO13X1', 'SL SO2X 1', 'SL EOLX 1', 
        'SL EOJX 1', 'SL EO12X1', 'SL EO13X1', 'SL EO2X 1', 'SL AOL 1', 'SL AOJ 1', 'SL AOJ121', 'SL AOJ131', 'SL AOJ2 1', 'SL ULC1 1', 'SL ULC1 2', 'SL ULC1 3', 
        'SL ULC2 1', 'SL ULC2 2', 'SL ULC2 3', 'SL ULC3 1', 'SL ULC3 2', 'SL ULC3 3','SL','UL','SL UL','SL UJ', 'SL AU','SL SE','SL SE', 'SL SU', 'SL SO', 'SL EO','SL AO','SUL2 YJ1021','SUL', 'SU']
        if lak.pla_search.get() in SunLtd:
            text1= ("Plan Group: ||Sun Ltd||SunSpectrum UL II||                                                                                ")
        elif lak.pla_search.get() in UL50:
            text1= ("Plan Group: ||Flexible Premium Life|| VIAPP||UL50||                                                                        ")
        elif lak.pla_search.get() in ULife_Met:
            text1= ("Plan Group: ||UL 2000||SunSpectrum UL||ULife Met||                                                                           ")
        elif lak.pla_search.get() in SunUL:
            text1= ("Plan Group: ||SunUL||SunUL Max||SunUL II||SunUL Pro||                                                                     ")
        else:
            text1=("Delete the last Letter and search again or refer to Knowledge Doc")
        # Label for Method
        Method_lbl= Label(lak.root, text=text1, font = ("Arial", 11, "bold"), fg= "white", bg= "#53085C")
        Method_lbl.place(x=10,y=410)
    def method(lak):
        if lak.var_plan.get() in ["SunUL", "SunUL Max","SunUL II", "SunUL Pro"]:
            step= "Amount= (Guaranted Cost of Insurance + Supplementry Benefit + Substandard Mortality Charges + RTB Charges)"
        else:
            step= "Amount= (Total Deduction) {:500} \n {:500})".format(" ", " ")
        # Label for Method
        Method_lbl= Label(lak.root, text=step, font = ("Arial", 12 ), fg= "white", bg= "#53085C")
        Method_lbl.place(x=10,y=440)
#_______________________________________________________________________________Second Screen______________________________________________________________________________________________________________
    def Withdrawal_Amount(lak):
        label_bg=Label(lak.root ,borderwidth=0, background="#0A0B0A")
        label_bg.place(x=0,y=0,relwidth=1,relheight=1)
        #hidden
        Suspense_lbl_entry =Entry(lak.root,width=45,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        Suspense_lbl_entry.place(x=200,y=580)
        Suspense_lbl_entry =Entry(lak.root, width=45,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        Suspense_lbl_entry.place(x=200,y=680)
        # Nevigation button
        Mbutton2= Button(lak.root, command= lak.Main_fram, activebackground= "#031b28",width=35, activeforeground= "white", 
        text = "COI Amount Caculation",fg= "black", bg= "white",font = ("Arial", 15, "bold "))
        Mbutton2.place(x=0,y=0)
        Mbutton2= Button(lak.root, command= lak.Withdrawal_Amount, activebackground= "#031b28",width=35, activeforeground= "white", 
        text = "Withdrawable/Calander",fg= "white", bg= "#0A0B0A",font = ("Arial", 15, "bold "))
        Mbutton2.place(x=400,y=0)
        titl_lbl= Label(lak.root,relief="raise" ,width= 80,justify="left", text= "Calander", font = ("Arial", 12, "bold"), fg= "white", bg= "#181B19")
        titl_lbl.place(x=0,y=47,height=40)
        lak.frame1= Label(lak.root,background="#0A0B0A", text= "Business Days / Regular Days :", font = ("Arial", 20, "bold"), fg= "white")
        lak.frame1.place (x=10,y=100)
        notes= Label(lak.root,text= " To find Lapse date => For MLIF : Select Previous Due Date \ For Ing : Select Lapse Start Date",
        justify="left",background="#0A0B0A", font = ("Arial",9, "bold"), fg= "white")
        notes.place(x=15,y=160,width=530,height= 40)
        PTD_lbl= Label(lak.root, text= "Select Date",width=15, font = ("Arial", 12, "bold"), fg= "white", bg= "#181B19")
        PTD_lbl.place(x=20,y=205,height=33)
        PTD_lbl_entry =Entry(lak.root,  width=30,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        PTD_lbl_entry.place(x=200,y=200)
        cal1= DateEntry(lak.root,selectmode="day",textvariable=lak.bday_date, borderwidth=0,foreground= "black",font = ("Arial",12, "bold"), width=28)
        cal1.place(x=207,y=210)
        No_of_days= Label(lak.root, text= "Number of Days",width=15, font = ("Arial", 12, "bold"), fg= "white", bg= "#181B19")
        No_of_days.place(x=20,y=255, height=33)
        No_of_days_entry =Entry(lak.root,textvariable=lak.number_days,  width=30,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        No_of_days_entry.place(x=200,y=250)
        Backbutton2= Button(lak.root, command=lak.Business_Days10 , activebackground= "yellow",activeforeground= "white", 
        text = "Business Days",fg= "white",width=15, bg= "#181B19",font = ("Arial", 12, "bold"))
        Backbutton2.place(x=20,y=305)
        but_lbl_10entry =Entry(lak.root,textvariable=lak.bday_date10,  width=30,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        but_lbl_10entry.place(x=200,y=300)
        Backbutton3= Button(lak.root, command=lak.Business_Days20, activebackground= "yellow",activeforeground= "white", 
        text = "Regular Days",fg= "white",width=15, bg= "#181B19",font = ("Arial", 12, "bold"))
        Backbutton3.place(x=20,y=355)
        but_lbl_20entry =Entry(lak.root,textvariable=lak.bday_date20,  width=30,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        but_lbl_20entry.place(x=200,y=350)
        Backbutton4= Button(lak.root, command=lak.Lapse_date, activebackground= "yellow",activeforeground= "white", 
        text = "Termination Date",fg= "white",width=15, bg= "#181B19",font = ("Arial", 12, "bold"))
        Backbutton4.place(x=20,y=405)
        but_lbl_4entry =Entry(lak.root,textvariable=lak.Lday_date62,  width=30,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        but_lbl_4entry.place(x=200,y=402)
        button= Button(lak.root, command=lak.Os_reset1, activebackground= "#031b28",activeforeground= "white", 
        text = "  Reset   ",fg= "white",width=47, bg= "#181B19",font = ("Arial", 12, "bold"))
        button.place(x=20,y=455)
        cal= DateEntry(lak.root,selectmode="day",textvariable=lak.w_hiddin_date, borderwidth=0,foreground= "black",font = ("Arial",12, "bold"), width=43)
        cal.place(x=559,y=90)
        Holidays=[]
        Datelist=[]
        F_Datelist=[]
        Eventlist=[]
        Year_is= int(datetime.now().strftime("%Y"))
        for day, name in sorted(holidays.Canada(years=Year_is).items()):
            hlist= (day, name)
            for i in hlist:
                if i not in Holidays:
                    Holidays.append(i)
        for j in range(0,len(Holidays),2):
            Datelist.append(Holidays[j])
        for K in range(1,len(Holidays),2):
            Eventlist.append(Holidays[K])
        for n in Datelist:
            F_Datelist.append(n.strftime("%d-%b"))
        S1= pd.Series(Eventlist, index= F_Datelist)
        Even_date_lbl= Label(lak.root,justify="left", text= S1,bg= "white", font = ("Arial",10), fg= "white",background="#0A0B0A")
        Even_date_lbl.place(x=580,y=300)
        Even_date_lbl= Label(lak.root,background="#0A0B0A")
        Even_date_lbl.place(x=540,y=495,width=300,height= 30)
        Listofholi= Label(lak.root,text= "List of Holidays(Canada)",justify="right",bg= "#0A0B0A", font = ("Arial",15, "bold"), fg= "white")
        Listofholi.place(x=533,y=270,width=300,height= 30)
        Calcyy= Calendar(lak.root,selectmode="day" )
        Calcyy.place(x=548,y=87)
        titl_lbl= Label(lak.root,relief="raise" ,width= 80,justify="left", text= "Withdrawable Amount", font = ("Arial", 12, "bold"), fg= "white", bg= "#181B19")
        titl_lbl.place(x=0,y=500,height=40)
        #Withdrawal
        Cprovince_lbl= Label(lak.root,justify="left", text= "Frequency",width= 15, borderwidth=10, font = ("Arial", 12, "bold"), fg= "white", bg= "#181B19")
        Cprovince_lbl.place(x=10,y=580)
        Que_combo= ttk.Combobox(lak.root, textvariable=lak.w_freq, width=48,font = ("Arial", 11, "bold"), state= "readonly")
        Que_combo["values"]=(" ","Anually","Monthly")
        Que_combo.set(" ")
        Que_combo.place(x=207,y=587)
        PTD_lbl= Label(lak.root,justify="left", text= "Policy Value",width= 15, borderwidth=10, font = ("Arial", 12, "bold"), fg= "white", bg= "#181B19")
        PTD_lbl.place(x=10,y=630)
        cal= Entry(lak.root,textvariable=lak.w_policy_val,  width=45,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        cal.place(x=200,y=630)
        NWD_lbl= Label(lak.root,justify="left", text= "Annivary date",width= 15, borderwidth=10, font = ("Arial", 12, "bold"), fg= "white", bg= "#181B19")
        NWD_lbl.place(x=10,y=680)
        cal= DateEntry(lak.root,selectmode="day",textvariable=lak.w_Anvercry_date, borderwidth=0,foreground= "black",font = ("Arial",12, "bold"), width=43)
        cal.place(x=207,y=687)
        Premium_lbl= Label(lak.root,justify="left", text= "COI",width= 15, borderwidth=10, font = ("Arial", 12, "bold"), fg= "white", bg= "#181B19")
        Premium_lbl.place(x=10,y=730)
        Premium_entry =Entry(lak.root,textvariable=lak.w_COI,  width=45,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        Premium_entry.place(x=200,y=730)
        Suspense= Label(lak.root,justify="left", text= "Amt. to be left",width= 15, borderwidth=10, font = ("Arial", 12, "bold"), fg= "white", bg= "#181B19")
        Suspense.place(x=10,y=780)
        Suspense_lbl_entry =Entry(lak.root,textvariable=lak.w_left,  width=45,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        Suspense_lbl_entry.place(x=200,y=780)
        Accumulated_lbl= Label(lak.root,justify="left", text= "Withdrawable Amt.",width= 15, borderwidth=10, font = ("Arial", 12, "bold"), fg= "white", bg= "#181B19")
        Accumulated_lbl.place(x=10,y=830)
        Suspense_lbl_entry =Entry(lak.root,textvariable=lak.w_with_amot,  width=45,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        Suspense_lbl_entry.place(x=200,y=830)
        button= Button(lak.root, command=lak.W_reset, activebackground= "#031b28",activeforeground= "white", 
        text = " Reset ",width=12,fg= "white", bg= "#181B19",font = ("Arial", 12, "bold"),borderwidth=5)
        button.place(x=650,y=680)
        button= Button(lak.root, command=lak.withdrawal, activebackground= "#031b28",activeforeground= "white", 
        text = " Enter ",width=12,fg= "white", bg= "#181B19",font = ("Arial", 12, "bold"),borderwidth=5)
        button.place(x=650,y=730)
        button4= Button(lak.root, command=lambda:pyperclip.copy(lak.w_left.get()), activebackground= "yellow",activeforeground= "white", 
        text = " Copy ",width=12,fg= "white", bg= "#181B19",font = ("Arial", 12, "bold"),borderwidth=5)
        button4.place(x=650,y=780)
        button5= Button(lak.root, command=lambda:pyperclip.copy(lak.w_with_amot.get()), activebackground= "yellow",activeforeground= "white", 
        text = " Copy ",width=12,fg= "white", bg= "#181B19",font = ("Arial", 12, "bold"),borderwidth=5)
        button5.place(x=650,y=830)
        
    def withdrawal(lak):
        P_date= (lak.w_hiddin_date.get()).split("/")
        N_date= (lak.w_Anvercry_date.get()).split("/")
        Month_os= int(N_date[0])-int(P_date[0])
        W_COI= float(lak.w_COI.get())
        W_Policy_V=float(lak.w_policy_val.get())
        if lak.w_freq.get()== "Anually":
            To_be_left= ((W_COI)*(Month_os))
            Withdrwalable_amount= (W_Policy_V-To_be_left)
            lak.w_left.set(To_be_left)
            lak.w_with_amot.set(Withdrwalable_amount)
        elif lak.w_freq.get()=="Monthly":
            To_be_left= W_COI
            Withdrwalable_amount= (W_Policy_V-To_be_left)
            lak.w_left.set(To_be_left)
            lak.w_with_amot.set(Withdrwalable_amount)

    def W_reset(lak):
        lak.w_freq.set("")
        lak.w_policy_val.set("")
        lak.w_Anvercry_date.set("")
        lak.w_COI.set("")
        lak.w_left.set("")
        lak.w_with_amot.set("")

        
    def Os_reset1(lak):
        lak.bday_date.set("")
        lak.bday_date10.set("")
        lak.bday_date20.set("")
        lak.Lday_date62.set("")   
        lak.number_days.set("")   
        return 

    def Business_Days10 (lak):
        number_of_days= int(lak.number_days.get())
        F_Datelist10=[]
        bizdate10= pd.date_range(start= lak.bday_date.get(),periods=number_of_days, freq="B")
        F_Datelist10.append(bizdate10.strftime("%d/%b/%Y"))
        bizdate10= df(F_Datelist10)
        B_ness_Days10=bizdate10.iloc[0].values[-1]
        lak.bday_date10.set(B_ness_Days10)

    def Business_Days20 (lak):
        number_of_days= int(lak.number_days.get())
        F_Datelist20=[]
        bizdate20= pd.date_range(start=lak.bday_date.get(),periods=number_of_days, freq="D")
        F_Datelist20.append(bizdate20.strftime("%d/%b/%Y"))
        bizdate20= df(F_Datelist20)
        B_ness_Days20=bizdate20.iloc[0].values[-1]
        lak.B_ness_Days20=bizdate20.values
        lak.bday_date20.set(B_ness_Days20)

    def Lapse_date (lak):
        F_Datelist62=[]
        bizdate62= pd.date_range(start=lak.bday_date.get(),periods=62, freq="D")
        F_Datelist62.append(bizdate62.strftime("%d/%b/%Y"))
        bizdate62= df(F_Datelist62)
        B_ness_Days62=bizdate62.iloc[0].values[-1]
        lak.Lday_date62.set(B_ness_Days62)
#--------------------------------------------------------------COI Function--------------------------------------------------------------------------------------
    def Os_reset(lak):
        lak.var_prems_amt.set("")
        lak.var_susp_amt.set("")
        lak.var_accum_amt.set("")
        lak.var_nsf_amt.set("")
        lak.var_outstand_amt.set("")
        lak.var_p_to_d.set("")
        lak.var_n_with_d.set("")
        lak.var_province2.set("")
        lak.var_plan2.set("")
        reset_outs_lbl= Label(lak.root, text=" ",width=400, font = ("Arial", 25), fg= "white", background="#53085C")
        reset_outs_lbl.place(x=10,y=900)
        return 

    def OutStanding(lak):
        province10 = ["British Columbia" , "Manitoba" , "New Brunswick", "Labrador", "Ontario", "Yukon","Saskatchewan(issued before 31 March,2000)"]
        province20 = ["Quebec"]
        province30= ["Newfoundland"]
        province40= ["Prince Edward Island"]
        province50= ["Alberta","Northwest Territories", "Nova Scotia" , "Nunavut", "Saskatchewan(issued after 31 March,2000)"]
        province60= ["Non Residence"]
        plan10= ["UL 2000", "SunSpectrum UL", "ULife Met","SunUL", "SunUL Max","SunUL II", "SunUL Pro"]
        plan20= ["Flexible Premium Life", "VIAPP", "UL50"]
        plan40= ["Sun Ltd","SunSpectrum UL II"]
        s0=lak.var_province2.get()
        premium= float(lak.var_prems_amt.get())
        Susp= float(lak.var_susp_amt.get())
        Accumilation= float(lak.var_accum_amt.get())
        NSF1= float(lak.var_nsf_amt.get())
        p2=lak.var_plan2.get()
        if p2 in plan10:
            if Accumilation < 0:
                if s0 in province10:
                    Accum= Accumilation/0.98
                elif s0 in province20:
                    Accum= (Accumilation/0.967)
                elif s0 in province30:
                    Accum= (Accumilation/0.96)
                elif s0 in province40:
                    Accum= (Accumilation/0.965)
                elif s0 in province50:
                    Accum= (Accumilation/0.97)
                elif s0 in province60:
                    Accum= (Accumilation/0.98)     
                else:
                    Accum= (Accumilation/0.98)
            else:
                Accum= Accumilation
        elif p2 in plan20:
            if Accumilation < 0:
                Accum= (Accumilation * 1.15)
            else:
                Accum= Accumilation
        elif p2 in plan40:
            if Accumilation < 0:
                Accum= (Accumilation/0.98)
        else:
            Accum= Accumilation
        P_date= (lak.var_p_to_d.get()).split("/")
        N_date= (lak.var_n_with_d.get()).split("/")
        Month_os= int(N_date[0])-int(P_date[0])
        outstanddingg= ((premium*Month_os) - Susp + NSF1 - Accum)
        lak.var_outstand_amt.set(round(outstanddingg,2))
        if outstanddingg<0:
            text2= ("There is no outstanding, policy have ${} in excess".format(-(outstanddingg)))
        else:
            text2=("Policy is outstanding for  ${:<500}".format(outstanddingg))
        # Label for Method
        outs_lbl= Label(lak.root, text=text2, font = ("Arial", 12,"bold"), fg= "white", background="#53085C")
        outs_lbl.place(x=10,y=900)
        pyperclip.copy(lak.var_outstand_amt.get())
        
    def Reset(lak):
        lak.var_province.set("")
        lak.var_plan.set("")
        lak.var_amount.set("")
        lak.pla_search.set("")
        lak.M_COI.set("")
        lak.A_COI.set("")
        # Label 
        Method_lbl= Label(lak.root, text=" ",width=400, borderwidth=0, font = ("Arial", 20, "bold"), fg= "black",bg= "#53085C")
        Method_lbl.place(x=0,y=400)
        Method_lbl= Label(lak.root, text=" ",width=400, font = ("Arial", 25), fg= "black", bg= "#53085C")
        Method_lbl.place(x=0,y=437)

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
    aap= Coi_tool(root)
    root.mainloop()