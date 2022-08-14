from tkinter import *
from tkinter import ttk
from tkinter import messagebox
import pyperclip

class Coi_Calci:
    
    def __init__(lak,root):
        lak.root= root
        lak.root.title("All In One Calulator")
        lak.root.geometry("600x700")
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
        lak.Main_fram()
#--------------------------------------------------------------First Screen--------------------------------------------------------------------------------------------------    
    
    def Main_fram(lak):
        #main Fram
        lak.frrame= Frame(lak.root,  bg= "white")
        lak.frrame.place (x=10,y=0,width=600,height=300)
        lak.frame= Frame(lak.root,  bg= "white")
        lak.frame.place (x=10,y=300,width=600,height=300)
        # button frame
        lak.frame2= Frame(lak.root,  bg= "white")
        lak.frame2.place (x=10,y=600,width=600,height=100)
        # Entrybox for search
        P_entry =Entry(lak.root,textvariable= lak.pla_search, width=30,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        P_entry.place(x=158,y=250)
        prod_lbl= Label(lak.root, text= "    Product Type",bg= "white", font = ("Arial", 12, "bold"), fg= "black")
        prod_lbl.place(x=10,y=250)
        # Label for Title
        titl_lbl= Label(lak.root, text= "Cost Of Insurance Amount:", font = ("Arial", 15, "bold"), fg= "black", bg= "white")
        titl_lbl.place(x=20,y=200)
        # Entrybox for province combo box
        COI_entry =Entry(lak.frame, width=30,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        COI_entry.place(x=150,y=50)
        # province combo box
        Que_combo= ttk.Combobox(lak.frame, textvariable=lak.var_province, width=31,font = ("Arial", 11, "bold"), state= "readonly")
        Que_combo["values"]=(" ","Non Residence","British Columbia" , "Manitoba" , "New Brunswick", "Labrador", "Ontario", "Yukon",
        "Alberta","Northwest Territories", "Nova Scotia" , "Nunavut","Quebec","Newfoundland","Prince Edward Island",
        "Saskatchewan(issued before 31 March,2000)","Saskatchewan(issued after 31 March,2000)")
        Que_combo.set(" ")
        Que_combo.place(x=160,y=60)
        #Label for province combo box
        Cprovince_lbl= Label(lak.frame, text= " Provinces",bg= "white", font = ("Arial", 12, "bold"), fg= "black")
        Cprovince_lbl.place(x=10,y=50)
        # Entrybox for plan combo box
        COI_entry =Entry(lak.frame, width=30,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        COI_entry.place(x=150,y=0)
        # Plan combo box
        Que_combo= ttk.Combobox(lak.frame, textvariable=lak.var_plan, width=31,font = ("Arial", 11, "bold"),state= "readonly")
        Que_combo["values"]= (" ","UL 2000", "SunSpectrum UL", "ULife Met","Flexible Premium Life", "VIAPP", "UL50","SunUL", "SunUL Max","SunUL II", 
        "SunUL Pro","Sun Ltd","SunSpectrum UL II")
        Que_combo.set(" ")
        Que_combo.place(x=160,y=10)
        #Label for plan combo box
        plan_lbl= Label(lak.frame, text= " Plan Type",bg= "white", font = ("Arial", 12, "bold"), fg= "black")
        plan_lbl.place(x=10,y=0)
        #Label for amount
        province_lbl= Label(lak.frame, text= " Amount   ",bg= "white", font = ("Arial", 12, "bold"), fg= "black")
        province_lbl.place(x=10,y=100)
        # Entrybox for amount
        Eamount =Entry(lak.frame,textvariable=lak.var_amount, width=25,foreground= "black",font = ("Arial", 14, "bold"),borderwidth=10)
        Eamount.place(x=150,y=100)
        # Label for Monthly COI
        COI_lbl= Label(lak.frame, text= "Monthly COI ",width= 10, borderwidth=10, font = ("Arial", 12, "bold"), fg= "black", bg= "white")
        COI_lbl.place(x=10,y=150)
        # Entrybox for Monthly COI
        COI_entry =Entry(lak.frame,textvariable=lak.M_COI, width=25,foreground= "black",font = ("Arial", 14, "bold"),borderwidth=10)
        COI_entry.place(x=150,y=150)
        # Label for Annually COI
        Annually_lbl= Label(lak.frame, text= "Annually COI ",width= 10, borderwidth=10, font = ("Arial", 12, "bold"), fg= "black", bg= "white")
        Annually_lbl.place(x=10,y=200)
        # Entrybox for Annually COI
        Annually_entry =Entry(lak.frame,textvariable=lak.A_COI,  width=25,borderwidth=10,foreground= "black",font = ("Arial", 14, "bold"))
        Annually_entry.place(x=150,y=200)
        # button
        button= Button(lak.root, command=lak.plan_search, activebackground= "#031b28",activeforeground= "white", 
        text = " Search ",fg= "black",width=8, bg= "white",font = ("Arial", 12, "bold"))
        button.place(x=490,y=253)
        button_method= Button(lak.frame, command=lak.method, activebackground= "#031b28",activeforeground= "white", 
        text = " Step ",fg= "black",width=8, bg= "white",font = ("Arial", 12, "bold"))
        button_method.place(x=480,y=0)
        button= Button(lak.frame, command=lak.Reset, activebackground= "#031b28",activeforeground= "white", 
        text = "  Reset   ",fg= "black",width=8, bg= "white",font = ("Arial", 12, "bold"))
        button.place(x=480,y=50)
        button= Button(lak.frame, command=lak.Coi_Calculator, activebackground= "#031b28",activeforeground= "white", 
        text = " Enter ",fg= "black",width=8, bg= "white",font = ("Arial", 12, "bold"))
        button.place(x=480,y=100)
        button4= Button(lak.frame, command=lambda:pyperclip.copy(lak.M_COI.get()), activebackground= "yellow",activeforeground= "white", 
        text = " Copy ",fg= "black",width=8, bg= "white",font = ("Arial", 12, "bold"))
        button4.place(x=480,y=150)
        button5= Button(lak.frame, command=lambda:pyperclip.copy(lak.A_COI.get()), activebackground= "yellow",activeforeground= "white", 
        text = " Copy ",fg= "black",width=8, bg= "white",font = ("Arial", 12, "bold"))
        button5.place(x=480,y=200)
        # Nevigation button
        Mbutton2= Button(lak.root, command= lak.Amount, activebackground= "#031b28",width=60, activeforeground= "white", 
        text = "   Amount Caculation   ",fg= "black", bg= "white",font = ("Arial", 12, "bold "))
        Mbutton2.place(x=0,y=0)
    def plan_search(lak):

        SunLtd=['SL Sun Ltd SLP101', 'SL Sun Ltd SLP102', 'SL Sun Ltd SLP151', 'SL Sun Ltd SLP152', 'SL Sun Ltd SLP201', 'SL Sun Ltd SLP202', 'SL Sun Ltd SLP651', 
        'SL Sun Ltd JLP101', 'SL Sun Ltd JLP102', 'SL Sun Ltd JLP151', 'SL Sun Ltd JLP201', 'SL Sun Ltd JLP202',"Sun Ltd","SunSpectrum UL II"]
        UL50=['Interest Plus', 'ZIP', 'Flexible Premium Life', 'UL50', 'Universal Life 50', 'VIAPP', 'Versatile Investment and Protection Plan']
        ULife_Met=['Universal Plus', 'Universal Plus joint', "Universal Plus (pre '92)", "Universal Plus joint (pre '92)Universal Flexiplus", 
        'Universal flexiplus joint', 'Universal Optimet', 'Universal Optimet joint', 'Interest Plus','SunSpectrum UL', 'SunSpectrum Universal Life', 'Jt. SunSpect UL', 
        'SunSpectrum Universal Life - Joint', 'SunSpectrum UL Increases', 
        'SunSpectrum Universal Life - Increase', 'Joint SunSpectrum UL Increases', 'SunSpectrum Universal Life - Joint Increase', 'Univ Life', 'Universal Life', 
        'UL - Splss', 'Universal Life - Special Issue', 'UL - SplssJt', 'Universal Life - Special Issue Joint', 'UL Inc (Universal Life - Increase)', 'UL Splss In', 
        'Universal Life - Special Issue Increase', 'UL Jt Inc', 'Universal Life - Joint Increase', 'Executive ULife',"UL 2000", "SunSpectrum UL", "ULife Met"]
        SunUL =["SunUL", "SunUL Max","SunUL II", "SunUL Pro",'SL ULA 1', 'SL ULA 2', 'SL ULA 3', 'SL ULA 4', 'SL ULA 5', 'SL ULA 6', 'SL ULA 7', 'SL ULA 8', 'SL ULA 9', 'SL ULJ 1', 'SL ULJ 2', 'SL ULJ 3', 
        'SL ULJ 4', 'SL ULJ 5', 'SL ULJ 6', 'SL ULJ 7', 'SL ULJ 8', 'SL ULJ 9', 'SL ULJ12 1', 'SL ULJ12 2', 'SL ULJ12 3', 'SL ULJ12 4', 'SL ULJ12 5', 'SL ULJ12 6', 
        'SL ULJ12 7', 'SL ULJ12 8', 'SL ULJ12 9', 'SL ULJ13 1', 'SL ULJ13 2', 'SL ULJ13 3', 'ULJ13 4', 'SL ULJ13 5', 'SL ULJ13 6', 'SL ULJ13 7', 'SL ULJ13 8', 
        'SL ULJ13 9', 'SL ULJ2 1', 'SL ULJ2 2', 'SL ULJ2 3', 'SL ULJ2 4', 'SL ULJ2 5', 'SL ULJ2 6', 'SL ULJ2 7', 'SL ULJ2 8', 'SL ULJ2 9', 'SL ULAX 1', 'SL ULAX 2', 
        'SL ULAX 3', 'SL ULAX 4', 'SL ULAX 5', 'SL ULAX 6', 'SL ULAX 7', 'SL ULAX 8', 'SL ULAX 9', 'SL SULX 1', 'SOLX 1', 'SL ULJX 1', 'ULJX 2', 'ULJX 3', 'ULJX 4', 
        'ULJX 5', 'ULJX 6', 'ULJX 7', 'ULJX 8', 'ULJX 9', 'SL UJ12X 1', 'UJ12X 2', 'UJ12X 3', 'UJ12X 4', 'UJ12X 5', 'UJ12X 6', 'UJ12X 7', 'UJ12X 8', 'UJ12X 9', 
        'SL UJ13X 1', 'UJ13X 2', 'UJ13X 3', 'UJ13X 4', 'UJ13X 5', 'UJ13X 6', 'UJ13X 7', 'UJ13X 8', 'UJ13X 9', 'SL UJ2X 1', 'UJ2X 2', 'UJ2X 3', 'UJ2X 4', 
        'UJ2X 5', 'UJ2X 6', 'UJ2X 7', 'UJ2X 8', 'UJ2X 9', 'SL ULC1 1', 'ULC1 2', 'ULC1 3', 'SL ULC1 1', 'ULC1 2', 'ULC1 3', 'SL ULC2 1', 'ULC2 2', 'ULC2 3', 
        'SL ULC3 1', 'ULC3 2', 'ULC3 3', 'SL SUL 1', 'SL SUJ 1', 'SL SUJ121', 'SL SUJ131', 'SL SUJ2 1', 'SL SEL 1', 'SL SEJ 1', 'SL SEJ121', 'SL SEJ131', 
        'SL SEJ2 1', 'SL SULX 1', 'SL SUJX 1', 'SL SU12X1', 'SL SU13X1', 'SL SU2X 1', 'SL SELX 1', 'SL SEJX 1', 'SL SE12X1', 'SL SE13X1', 'SL SE2X 1', 'SL AUL 1', 
        'SL AUJ 1', 'SL AUJ121', 'SL AUJ131', 'SL AUJ2 1', 'SL ULA A', 'ULA B', 'ULA C', 'ULA D', 'ULA E', 'ULA F', 'ULA G', 'ULA H', 'SL ULJ A', 'ULJ B', 'ULJ C', 
        'ULJ D', 'ULJ E', 'ULJ F', 'ULJ G', 'ULJ H', 'SL ULJ12 A', 'ULJ12 B', 'ULJ12 C', 'ULJ12 D', 'ULJ12 E', 'ULJ12 F', 'ULJ12 G', 'ULJ12 H', 'SL ULJ13 A', 
        'ULJ13 B', 'ULJ13 C', 'ULJ13 D', 'ULJ13 E', 'ULJ13 F', 'ULJ13 G', 'ULJ13 H', 'SL ULJ2 A', 'ULJ2 B', 'ULJ2 C', 'ULJ2 D', 'ULJ2 E', 'ULJ2 F', 'ULJ2 G', 
        'ULJ2 H', 'SL SOL 1', 'SL SOJ 1', 'SL SOJ121', 'SL SOJ131', 'SL SOJ2 1', 'SL EOL 1', 'SL EOJ 1', 'SL EOJ121', 'SL EOJ131', 'SL EOJ2 1', 'SL SOLX 1', 
        'SL SOJX 1', 'SL SO12X1', 'SL SO13X1', 'SL SO2X 1', 'SL EOLX 1', 'SL EOJX 1', 'SL EO12X1', 'SL EO13X1', 'SL EO2X 1', 'SL AOL 1', 'SL AOJ 1', 'SL AOJ121', 
        'SL AOJ131', 'SL AOJ2 1', 'SL ULC1 1', 'ULC1 2', 'ULC1 3', 'SL ULC2 1', 'ULC2 2', 'ULC2 3', 'SL ULC3 1', 'ULC3 2', 'ULC3 3'
        'SL UJ12X 2', 'SL UJ12X 3', 'SL UJ12X 4', 'SL UJ12X 5', 'SL UJ12X 6', 'SL UJ12X 7', 'SL UJ12X 8', 'SL UJ12X 9', 
        'SL UJ13X 1', 'SL UJ13X 2', 'SL UJ13X 3', 'SL UJ13X 4', 'SL UJ13X 5', 'SL UJ13X 6', 'SL UJ13X 7', 'SL UJ13X 8', 'SL UJ13X 9', 'SL UJ2X 1', 'SL UJ2X 2', 
        'SL UJ2X 3', 'SL UJ2X 4', 'SL UJ2X 5', 'SL UJ2X 6', 'SL UJ2X 7', 'SL UJ2X 8', 'SL UJ2X 9', 'SL ULC1 1', 'SL ULC1 2', 'SL ULC1 3', 'SL ULC1 1', 'SL ULC1 2', 
        'SL ULC1 3', 'SL ULC2 1', 'SL ULC2 2', 'SL ULC2 3', 'SL ULC3 1', 'SL ULC3 2', 'SL ULC3 3', 'SL SUL 1', 'SL SUJ 1', 'SL SUJ121', 'SL SUJ131', 'SL SUJ2 1', 'SL SEL 1', 
        'SL SEJ 1', 'SL SEJ121', 'SL SEJ131', 'SL SEJ2 1', 'SL SULX 1', 'SL SUJX 1', 'SL SU12X1', 'SL SU13X1', 'SL SU2X 1', 'SL SELX 1', 
        'SL SEJX 1', 'SL SE12X1', 'SL SE13X1', 'SL SE2X 1', 'SL AUL 1', 'SL AUJ 1', 'SL AUJ121', 'SL AUJ131', 
        'SL AUJ2 1', 'SL ULA A', 'SL ULA B', 'SL ULA C', 'SL ULA D', 'SL ULA E', 'SL ULA F', 'SL ULA G', 'SL ULA H', 'SL ULJ A', 'SL ULJ B', 
        'SL ULJ C', 'SL ULJ D', 'SL ULJ E', 'SL ULJ F', 'SL ULJ G', 'SL ULJ H', 'SL ULJ12 A', 'SL ULJ12 B', 'SL ULJ12 C', 'SL ULJ12 D', 'SL ULJ12 E', 
        'SL ULJ12 F', 'SL ULJ12 G', 'SL ULJ12 H', 'SL ULJ13 A', 'SL ULJ13 B', 'SL ULJ13 C', 'SL ULJ13 D', 'SL ULJ13 E', 'SL ULJ13 F', 'SL ULJ13 G', 
        'SL ULJ13 H', 'SL ULJ2 A', 'SL ULJ2 B', 'SL ULJ2 C', 'SL ULJ2 D', 'SL ULJ2 E', 'SL ULJ2 F', 'SL ULJ2 G', 'SL ULJ2 H', 'SL SOL 1', 'SL SOJ 1', 
        'SL SOJ121', 'SL SOJ131', 'SL SOJ2 1', 'SL EOL 1', 'SL EOJ 1', 'SL EOJ121', 'SL EOJ131', 'SL EOJ2 1', 'SL SOLX 1', 'SL SOJX 1', 
        'SL SO12X1', 'SL SO13X1', 'SL SO2X 1', 'SL EOLX 1', 'SL EOJX 1', 'SL EO12X1', 'SL EO13X1', 'SL EO2X 1', 'SL AOL 1', 'SL AOJ 1', 
        'SL AOJ121', 'SL AOJ131', 'SL AOJ2 1', 'SL ULC1 1', 'SL ULC1 2', 'SL ULC1 3', 'SL SL ULC2 1', 'SL ULC2 2', 'SL ULC2 3', 'SL ULC3 1', 'SL ULC3 2', 'SL ULC3 3'
        'SL','UL','SL UL','SL UJ', 'SL AU','SL SE','SL SE', 'SL SU', 'SL SO', 'SL EO','SL AO']
        if lak.pla_search.get() in SunLtd:
            text1= ("Plan Group: ||Sun Ltd||SunSpectrum UL II||")
        elif lak.pla_search.get() in UL50:
            text1= ("Plan Group: ||Flexible Premium Life|| VIAPP||UL50||")
        elif lak.pla_search.get() in ULife_Met:
            text1= ("Plan Group: ||UL 2000||SunSpectrum UL||ULife Met||")
        elif lak.pla_search.get() in SunUL:
            text1= ("Plan Group: ||SunUL||SunUL Max||SunUL II||SunUL Pro||")
        else:
            text1=("Delete the last Letter and search again or refer to Knowledge Doc")
        # Label for Method
        Method_lbl= Label(lak.root,width=50, text=text1, font = ("Arial", 12, "bold"), fg= "black", bg= "white")
        Method_lbl.place(x=20,y=570)

#--------------------------------------------------------------Second Screen--------------------------------------------------------------------------------------
    def method(lak):
        if lak.var_plan.get() in ["SunUL", "SunUL Max","SunUL II", "SunUL Pro"]:
            step= """Amount= (Guaranted Cost of Insurance + Supplementry Benefit Charges + 
            Substandard Mortality Charges + RTB Charges)                                """
        else:
            step= "Amount= (Total Deduction) {:500} \n {:500})".format(" ", " ")
        # Label for Method
        Method_lbl= Label(lak.root, text=step, font = ("Arial", 12, "bold"), fg= "black", bg= "white")
        Method_lbl.place(x=20,y=600)

    def Amount(lak):
        lak.Mframe= Frame(lak.root,  bg= "white")
        lak.Mframe.place (x=0,y=0,width=600,height=800)
        # Nevigation button
        Mbutton2= Button(lak.root, command= lak.Main_fram, activebackground= "#031b28",width=60, activeforeground= "white", 
        text = "   COI Caculation   ",fg= "black", bg= "white",font = ("Arial", 12, "bold "))
        Mbutton2.place(x=0,y=0)
        entry = Entry(lak.Mframe, relief=SUNKEN,borderwidth=10,foreground= "black",font = ("Arial", 25, "bold"), width=31)
        entry.place(x=13,y=34)
        COI_entry =Entry(lak.Mframe, width=30,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        COI_entry.place(x=200,y=200)
        Que_combo= ttk.Combobox(lak.Mframe, textvariable=lak.var_province2, width=31,font = ("Arial", 11, "bold"), state= "readonly")
        Que_combo["values"]=(" ","Non Residence","British Columbia" , "Manitoba" , "New Brunswick", "Labrador", "Ontario", "Yukon",
        "Alberta","Northwest Territories", "Nova Scotia" , "Nunavut","Quebec","Newfoundland","Prince Edward Island",
        "Saskatchewan(issued before 31 March,2000)","Saskatchewan(issued after 31 March,2000)")
        Que_combo.set(" ")
        Que_combo.place(x=210,y=208)
        #Label for province combo box
        Cprovince_lbl= Label(lak.Mframe, text= "Provinces                ",bg= "white",width= 15,font = ("Arial", 12, "bold"), fg= "black")
        Cprovince_lbl.place(x=10,y=207)
        # Label for Paid to Dat
        PTD_lbl= Label(lak.Mframe, text= "Paid to date           ",width= 15, borderwidth=10, font = ("Arial", 12, "bold"), fg= "black", bg= "white")
        PTD_lbl.place(x=0,y=250)
        # Entrybox for Paid to date
        PTD_lbl_entry =Entry(lak.Mframe,textvariable=lak.var_p_to_d, width=30,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        PTD_lbl_entry.place(x=200,y=250)
        # Label for Next Withdrawal date
        NWD_lbl= Label(lak.Mframe, text= "Next Withdrawal   ",width= 15, font = ("Arial", 12, "bold"), fg= "black", bg= "white")
        NWD_lbl.place(x=10,y=300)
        # Entrybox for Next Withdrawal Date
        NWD_entry =Entry(lak.Mframe,textvariable=lak.var_n_with_d,  width=30,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        NWD_entry.place(x=200,y=300)
        plan_lbl1= Label(lak.Mframe, text= " Plan Type               ",bg= "white", font = ("Arial", 12, "bold"), fg= "black")
        plan_lbl1.place(x=10,y=350)
        # Entrybox for Paid to date
        PTD_lbl_entry =Entry(lak.Mframe,  width=30,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        PTD_lbl_entry.place(x=200,y=350)
        # Plan combo box
        Que_combo2= ttk.Combobox(lak.Mframe, textvariable=lak.var_plan2, width=31,font = ("Arial", 11, "bold"),state= "readonly")
        Que_combo2["values"]= (" ","UL 2000", "SunSpectrum UL", "ULife Met","Flexible Premium Life", "VIAPP", "UL50","SunUL", "SunUL Max","SunUL II", 
        "SunUL Pro","Sun Ltd","SunSpectrum UL II")
        Que_combo2.set(" ")
        Que_combo2.place(x=210,y=360)
        # Label for Premium
        Premium_lbl= Label(lak.Mframe, text= " Min. Premium Amt. ",width= 15, font = ("Arial", 12, "bold"), fg= "black", bg= "white")
        Premium_lbl.place(x=10,y=400)
        # Entrybox for Premium
        Premium_entry =Entry(lak.Mframe,textvariable=lak.var_prems_amt,  width=30,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        Premium_entry.place(x=200,y=400)
        # Label for Suspense
        Suspense= Label(lak.Mframe, text= "Suspense Amt.       ",width= 15, font = ("Arial", 12, "bold"), fg= "black", bg= "white")
        Suspense.place(x=10,y=450)
        # Entrybox for Suspense
        Suspense_lbl_entry =Entry(lak.Mframe,textvariable=lak.var_susp_amt,  width=30,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        Suspense_lbl_entry.place(x=200,y=450)
        # Label for Accumulated
        Accumulated_lbl= Label(lak.Mframe, text= "Accumulated Amt.  ",width= 15, font = ("Arial", 12, "bold"), fg= "black", bg= "white")
        Accumulated_lbl.place(x=10,y=500)
        # Entrybox for Accumulated
        Accumulated_lbl_entry =Entry(lak.Mframe,textvariable=lak.var_accum_amt,  width=30,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        Accumulated_lbl_entry.place(x=200,y=500)
        # Label for NSF Fee
        Fee_lbl= Label(lak.Mframe, text= "NSF Fee                  ",width= 15, font = ("Arial", 12, "bold"), fg= "black", bg= "white")
        Fee_lbl.place(x=10,y=550)
        # Entrybox for NSF FEE
        Fee_lbl_entry =Entry(lak.Mframe,textvariable=lak.var_nsf_amt,  width=30,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        Fee_lbl_entry.place(x=200,y=550)
        # Label for outstanding
        outs_lbl= Label(lak.Mframe, text= "Outstanding Amt. ",width= 15, font = ("Arial", 12, "bold"), fg= "black", bg= "white")
        outs_lbl.place(x=10,y=600)
        # Entrybox for outstandig
        outs_lbl_entry =Entry(lak.Mframe,textvariable=lak.var_outstand_amt,  width=30,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        outs_lbl_entry.place(x=200,y=600)
        #Copy for Calculator
        Copybutton= Button(lak.Mframe, command=lak.OutStanding, activebackground= "yellow",activeforeground= "white", 
        text = " OutStanding ",fg= "black",width=19, bg= "white",font = ("Arial", 12, "bold"))
        Copybutton.place(x=400,y=90)
        #Copy for outstanding
        Copy2button= Button(lak.Mframe, command=lambda:pyperclip.copy(lak.var_outstand_amt.get()), activebackground= "yellow",activeforeground= "white", 
        text = " Copy ",fg= "black",width=8, bg= "white",font = ("Arial", 12, "bold"))
        Copy2button.place(x=500,y=600)
        # Label for osamt
        PTD_lbl= Label(lak.Mframe, text= "Find Outstanding Amount:", font = ("Arial", 12, "bold"), fg= "black", bg= "white")
        PTD_lbl.place(x=15,y=140)
        button= Button(lak.Mframe, command=lak.Os_reset, activebackground= "#031b28",activeforeground= "white", 
        text = "  Reset   ",fg= "black",width=8, bg= "white",font = ("Arial", 12, "bold"))
        button.place(x=500,y=550)
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
        button_clear = Button(master=lak.Mframe, text="Clear", command=clear,width=19,activebackground= "#031b28", activeforeground= "white",
        fg= "black", bg= "white",font = ("Arial", 12, "bold"))
        button_clear.place(x=0,y=90)
        button_equal = Button(master=lak.Mframe, text="Equal",command=equal,width=19,activebackground= "#031b28", activeforeground= "white",
        fg= "black", bg= "white",font = ("Arial", 12, "bold"))
        button_equal.place(x=200,y=90)

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
        reset_outs_lbl= Label(lak.root, text=" ",width=400, font = ("Arial", 25), fg= "black", bg= "white")
        reset_outs_lbl.place(x=10,y=650)
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
            if Accumilation < float(lak.var_prems_amt.get()):
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
        Month_os= int(N_date[1])-int(P_date[1])
        outstanddingg= ((premium*Month_os) - Susp + NSF1 - Accum)
        lak.var_outstand_amt.set(round(outstanddingg,2))
        if outstanddingg<0:
            text2= ("There is no outstanding, policy have ${} in excess".format(-(outstanddingg)))
        else:
            text2=("Policy is outstanding for  ${:<500}".format(outstanddingg))
        # Label for Method
        outs_lbl= Label(lak.root, text=text2, font = ("Arial", 12,"bold"), fg= "black", bg= "white")
        outs_lbl.place(x=10,y=650)
        
    def Reset(lak):
        lak.var_province.set("")
        lak.var_plan.set("")
        lak.var_amount.set("")
        lak.pla_search.set("")
        lak.M_COI.set("")
        lak.A_COI.set("")
        # Label 
        Method_lbl= Label(lak.root, text=" ",width=400, borderwidth=0, font = ("Arial", 20, "bold"), fg= "black",bg= "white")
        Method_lbl.place(x=20,y=570)
        Method_lbl= Label(lak.root, text=" ",width=400, font = ("Arial", 25), fg= "black", bg= "white")
        Method_lbl.place(x=20,y=600)        
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
    



