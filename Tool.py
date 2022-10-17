from tkinter import *
from tkinter import ttk
from tkcalendar import *
from tkcalendar import DateEntry
from tkinter import messagebox
from PIL import Image,ImageTk
import pyperclip
import pandas as pd
from pandas import DataFrame as df
import holidays
from datetime import datetime
from pandas.tseries.offsets import CustomBusinessDay
import webbrowser

class Coi_Calci:
    
    def __init__(lak,root):
        lak.root= root
        lak.root.title("All In One Calulator")
        lak.root.geometry("600x800")
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
        lak.Main_fram()
#--------------------------------------------------------------Calendar Screen--------------------------------------------------------------------------------------
    def Calander(lak):
        # backgroung image
        lak.bg=ImageTk.PhotoImage(file= "C:\\Users\\j683\\Desktop\\py\\slbg.jpg")
        label_cal=Label(lak.root, image=lak.bg,borderwidth=0)
        label_cal.place(x=0,y=0,relwidth=1,relheight=1)
        Calcyy= Calendar(label_cal,selectmode="day" )
        Calcyy.place(x=290,y=400, width=300,height= 250)
        # Nevigation button
        Mbutton2= Button(lak.root, command= lak.Main_fram, activebackground= "#031b28",width=19, activeforeground= "white", 
        text = "   COI Caculation   ",fg= "black", bg= "white",font = ("Arial", 12, "bold "))
        Mbutton2.place(x=0,y=0)
        Mbutton2= Button(lak.root, command= lak.GUI, activebackground= "#031b28",width=19, activeforeground= "white", 
        text = " Address ",fg= "black", bg= "white",font = ("Arial", 12, "bold"))
        Mbutton2.place(x=200,y=0)
        Mbutton2= Button(lak.root, command= lak.Amount, activebackground= "#031b28",width=19, activeforeground= "white", 
        text = " Amount Caculation ",fg= "black", bg= "white",font = ("Arial", 12, "bold"))
        Mbutton2.place(x=400,y=0)
        #Back Button
        button= Button(lak.root, command=lak.Os_reset1, activebackground= "#031b28",activeforeground= "white", 
        text = "  Reset   ",fg= "black",width=13, bg= "white",font = ("Arial", 12, "bold"))
        button.place(x=450,y=650)
        # Label for Paid to Dat
        notes= Label(lak.root,text= " To find Lapse date => For MLIF : Select Previous Due Date \ For Ing : Select Lapse Start Date",
        justify="left",bg= "white", font = ("Arial",9, "bold"), fg= "black")
        notes.place(x=15,y=342,width=530,height= 40)
        PTD_lbl= Label(lak.root, text= "Select Date", borderwidth=10, font = ("Arial", 12, "bold"), fg= "black", bg= "white")
        PTD_lbl.place(x=10,y=270)
        # Entrybox for Paid to date
        PTD_lbl_entry =Entry(lak.root,  width=25,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        PTD_lbl_entry.place(x=20,y=310)
        #DATE
        cal1= DateEntry(lak.root,selectmode="day",textvariable=lak.bday_date, borderwidth=0,foreground= "black",font = ("Arial",12, "bold"), width=23)
        cal1.place(x=27,y=317)

        No_of_days= Label(lak.root, text= "Number of Days", borderwidth=10, font = ("Arial", 12, "bold"), fg= "black", bg= "white")
        No_of_days.place(x=10,y=370)
        # Entrybox for Paid to date
        No_of_days_entry =Entry(lak.root,textvariable=lak.number_days,  width=25,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        No_of_days_entry.place(x=20,y=405)

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
        Even_date_lbl= Label(lak.root,justify="left", text= S1,bg= "white", font = ("Arial",10), fg= "black")
        Even_date_lbl.place(x=0,y=55,width=290,height= 250)
        PTD_lbl= Label(lak.root, text= "Select Date", borderwidth=10, font = ("Arial", 12, "bold"), fg= "black", bg= "white")
        PTD_lbl.place(x=10,y=269)
        Listofholi= Label(lak.root,text= "List of Holidays(Canada)",justify="right",bg= "white", font = ("Arial",15, "bold"), fg= "black")
        Listofholi.place(x=0,y=40,width=300,height= 50)
        Backbutton2= Button(lak.root, command=lak.Business_Days10 , activebackground= "yellow",activeforeground= "white", 
        text = "Business Days",fg= "black",width=23, bg= "white",font = ("Arial", 12, "bold"))
        Backbutton2.place(x=20,y=450)
        but_lbl_10entry =Entry(lak.root,textvariable=lak.bday_date10,  width=25,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        but_lbl_10entry.place(x=20,y=490)
        Backbutton3= Button(lak.root, command=lak.Business_Days20, activebackground= "yellow",activeforeground= "white", 
        text = "Regular Days",fg= "black",width=23, bg= "white",font = ("Arial", 12, "bold"))
        Backbutton3.place(x=20,y=530)
        but_lbl_20entry =Entry(lak.root,textvariable=lak.bday_date20,  width=25,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        but_lbl_20entry.place(x=20,y=570)
        Backbutton4= Button(lak.root, command=lak.Lapse_date, activebackground= "yellow",activeforeground= "white", 
        text = "Policy Lapse Date",fg= "black",width=23, bg= "white",font = ("Arial", 12, "bold"))
        Backbutton4.place(x=20,y=610)
        but_lbl_4entry =Entry(lak.root,textvariable=lak.Lday_date62,  width=25,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        but_lbl_4entry.place(x=20,y=650)
        
    def Os_reset1(lak):
        lak.var_prems_amt.set("")
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

#--------------------------------------------------------------First Screen--------------------------------------------------------------------------------------------------    
    
    def Main_fram(lak):
        # backgroung image
        lak.bg=ImageTk.PhotoImage(file="C:\\Users\\j683\\Desktop\\py\\slbg.jpg")
        label_bg=Label(lak.root, image=lak.bg,borderwidth=0)
        label_bg.place(x=0,y=0,relwidth=1,relheight=1)
        #main Fram
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
        Mbutton2= Button(lak.root, command= lak.Amount, activebackground= "#031b28",width=19, activeforeground= "white", 
        text = "   Amount Caculation   ",fg= "black", bg= "white",font = ("Arial", 12, "bold "))
        Mbutton2.place(x=0,y=0)
        Mbutton2= Button(lak.root, command= lak.GUI, activebackground= "#031b28",width=19, activeforeground= "white", 
        text = " Address ",fg= "black", bg= "white",font = ("Arial", 12, "bold"))
        Mbutton2.place(x=200,y=0)
        Mbutton2= Button(lak.root, command= lak.Calander, activebackground= "#031b28",width=19, activeforeground= "white", 
        text = " Calender ",fg= "black", bg= "white",font = ("Arial", 12, "bold"))
        Mbutton2.place(x=400,y=0)


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
        lak.fbg=ImageTk.PhotoImage(file="C:\\Users\\j683\\Desktop\\py\\slbg.jpg")
        label_fbg=Label(lak.Mframe, image=lak.fbg,borderwidth=0)
        label_fbg.place(x=0,y=0,relwidth=1,relheight=1)
        # Nevigation button
        Mbutton2= Button(lak.root, command= lak.Main_fram, activebackground= "#031b28",width=19, activeforeground= "white", 
        text = "   COI Caculation   ",fg= "black", bg= "white",font = ("Arial", 12, "bold "))
        Mbutton2.place(x=0,y=0)
        Mbutton2= Button(lak.root, command= lak.GUI, activebackground= "#031b28",width=19, activeforeground= "white", 
        text = " Address ",fg= "black", bg= "white",font = ("Arial", 12, "bold"))
        Mbutton2.place(x=200,y=0)
        Mbutton2= Button(lak.root, command= lak.Calander, activebackground= "#031b28",width=19, activeforeground= "white", 
        text = " Calender ",fg= "black", bg= "white",font = ("Arial", 12, "bold"))
        Mbutton2.place(x=400,y=0)
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
        PTD_lbl_entry =Entry(lak.Mframe,  width=30,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        PTD_lbl_entry.place(x=200,y=250)
        # Label for Next Withdrawal date
        NWD_lbl= Label(lak.Mframe, text= "Next Withdrawal   ",width= 15, font = ("Arial", 12, "bold"), fg= "black", bg= "white")
        NWD_lbl.place(x=10,y=300)
        # Entrybox for Next Withdrawal Date
        NWD_entry =Entry(lak.Mframe,  width=30,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        NWD_entry.place(x=200,y=300)
        #DATE
        cal= DateEntry(lak.Mframe,selectmode="day",textvariable=lak.var_p_to_d, borderwidth=0,foreground= "black",font = ("Arial",12, "bold"), width=28)
        cal.place(x=210,y=258)
        cal= DateEntry(lak.Mframe,selectmode="day",textvariable=lak.var_n_with_d, borderwidth=0,foreground= "black",font = ("Arial",12, "bold"), width=28)
        cal.place(x=210,y=308)
        #Label for plan combo box
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
        Month_os= int(N_date[0])-int(P_date[0])
        outstanddingg= ((premium*Month_os) - Susp + NSF1 - Accum)
        lak.var_outstand_amt.set(outstanddingg)
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

#--------------------------------------------------------------Address Screen--------------------------------------------------------------------------------------
    def GUI(lak):
        # backgroung image
        lak.bg=ImageTk.PhotoImage(file="C:\\Users\\j683\\Desktop\\py\\slbg.jpg")
        label_bg=Label(lak.root, image=lak.bg,borderwidth=0)
        label_bg.place(x=0,y=0,relwidth=1,relheight=1)
        # Nevigation button
        Mbutton2= Button(lak.root, command= lak.Main_fram, activebackground= "#031b28",width=19, activeforeground= "white", 
        text = "   COI Caculation   ",fg= "black", bg= "white",font = ("Arial", 12, "bold "))
        Mbutton2.place(x=0,y=0)
        Mbutton2= Button(lak.root, command= lak.Calander, activebackground= "#031b28",width=19, activeforeground= "white", 
        text = " Calender ",fg= "black", bg= "white",font = ("Arial", 12, "bold"))
        Mbutton2.place(x=200,y=0)
        Mbutton2= Button(lak.root, command= lak.Amount, activebackground= "#031b28",width=19, activeforeground= "white", 
        text = " Amount Caculation ",fg= "black", bg= "white",font = ("Arial", 12, "bold"))
        Mbutton2.place(x=400,y=0)
        # frame for Addreess entry box
        lak.frame= Frame(lak.root,  bg= "white")
        lak.frame.place (x=20,y=300,width=600,height=400)
        lak.frame1= Label(lak.root,  bg= "white", text= "Enter Address:", font = ("Arial", 12, "bold"), fg= "black")
        lak.frame1.place (x=30,y=270)
        # Entrybox for Addreess
        entry =Entry(lak.frame, width=46,textvariable=lak.Var_address, borderwidth=10,foreground= "black",font = ("Arial", 15, "bold"))
        entry.place(x=15,y=1,height=50)
        #Split Button
        button1= Button(lak.frame, command=lak.Split, activebackground= "yellow",activeforeground= "white", 
        text = " Split ",fg= "black",width=17, bg= "white",font = ("Arial", 12, "bold"))
        button1.place(x=15,y=50)
        #Reset Button
        button1= Button(lak.frame, command=lak.reset, activebackground= "yellow",activeforeground= "white", 
        text = " Reset ",fg= "black",width=16, bg= "white",font = ("Arial", 12, "bold"))
        button1.place(x=366,y=50)
        #Reset Button
        button1= Button(lak.frame, command=lambda:pyperclip.copy(lak.address_to_copy), activebackground= "yellow",activeforeground= "white", 
        text = " Copy ",fg= "black",width=16, bg= "white",font = ("Arial", 12, "bold"))
        button1.place(x=195,y=50)
        #Search
        Copybutton= Button(lak.root,command=lak.callback,  activebackground= "yellow",activeforeground= "white", 
        text = " Google Search ",fg= "black",width=27, bg= "white",font = ("Arial", 12, "bold"))
        Copybutton.place(x=20,y=670)
        Copybutton= Button(lak.root,command=lak.callback_canada,  activebackground= "yellow",activeforeground= "white", 
        text = " Canada Post",fg= "black",width=27, bg= "white",font = ("Arial", 12, "bold"))
        Copybutton.place(x=300,y=670)
        #Label for Address line 1
        Line= Label(lak.frame, text= "Line1 :",bg= "white", font = ("Arial", 12, "bold"), fg= "black")
        Line.place(x=10,y=100)
        # Entrybox for Address line 1
        Line1 =Entry(lak.frame,textvariable= lak.address1, width=35,foreground= "black",font = ("Arial", 14, "bold"),borderwidth=10)
        Line1.place(x=120,y=100)
        #Label for Address line 2
        Line= Label(lak.frame, text= "Line2 :",bg= "white", font = ("Arial", 12, "bold"), fg= "black")
        Line.place(x=10,y=150)
        # Entrybox for Address line 2
        lak.Line2 =Entry(lak.frame,textvariable= lak.address2, width=35,foreground= "black",font = ("Arial", 14, "bold"),borderwidth=10)
        lak.Line2.place(x=120,y=150)
        #Label for City
        Zip_code= Label(lak.frame, text= "City :",bg= "white", font = ("Arial", 12, "bold"), fg= "black")
        Zip_code.place(x=10,y=200)
        # Entrybox for City
        Zip_code =Entry(lak.frame,textvariable= lak.city , width=35,foreground= "black",font = ("Arial", 14, "bold"),borderwidth=10)
        Zip_code.place(x=120,y=200)
        #Label for Province
        Province= Label(lak.frame, text= "Province :",bg= "white", font = ("Arial", 12, "bold"), fg= "black")
        Province.place(x=10,y=250)
        # Entrybox for Province
        Province =Entry(lak.frame,textvariable= lak.Province, width=35,foreground= "black",font = ("Arial", 14, "bold"),borderwidth=10)
        Province.place(x=120,y=250)
        #Label for Zip_code
        Zip_code= Label(lak.frame, text= "Zip_code :",bg= "white", font = ("Arial", 12, "bold"), fg= "black")
        Zip_code.place(x=10,y=300)
        # Entrybox for Zip_code
        Zip_code =Entry(lak.frame,textvariable= lak.Zip_code , width=35,foreground= "black",font = ("Arial", 14, "bold"),borderwidth=10)
        Zip_code.place(x=120,y=300)

#--------------------------------------------------------------Address Function--------------------------------------------------------------------------------------
    def callback_canada(lak):
        webbrowser.open_new(r"https://www.canadapost-postescanada.ca/cpc/en/home.page")

    def callback(lak):
        webbrowser.open_new(r"https://www.google.co.in/")
        

    def reset(lak):
        lak.Var_address.set("")
        lak.address1.set("")
        lak.address2.set("")
        lak.city.set("")
        lak.Province.set("")
        lak.Zip_code.set("")
        return
    def Split(lak):
        lak.Address= lak.Var_address.get()
        lak.Address=(lak.Address.split(" "))
        #for Zip
        z= str(lak.Address[-1])
        n= len(z)
        if n==6:
            lak.pin=z
        if n==3:
            lak.pin= str(lak.Address[-2])+" "+ str(lak.Address[-1])
        lak.pin= lak.pin.split(" ")
        address= []
        for i in range(len(lak.Address)):
            if lak.Address[i] not in lak.pin:
                address.append(lak.Address[i])
        pin= " ".join(lak.pin)
        #for Province
        last_p1= ["Columbia", "COLUMBIA","Brunswick", "BRUNSWICK","Territories", "TERRITORIES","SCOTIA","Scotia",
        "Virginia", "VIRGINIA", "Dakota", "DAKOTA","Jersey", "JERSEY","Mexico","MEXICO","HEMISPHERE","Hemisphere"]
        last_p2= ["Labrador","ISLAND","Island","LABRADOR"]
        short_p= {"Alberta": "AB" , 
        "British Columbia" : "BC", 
        "Manitoba" :"MB", 
        "New Brunswick": "NB", 
        "Newfoundland and Labrador" :"NL", 
        "Northwest Territories": "NT", 
        "Nova Scotia": "NS", 
        "Nunavut": "NU", 
        "Ontario": "ON", 
        "Prince Edward Island": "PE", 
        "Quebec":"QC" , 
        "Saskatchewan": "SK" , 
        "Yukon": "YT",
        "ALBERTA": "AB", 
        "BRITISH COLUMBIA": "BC", 
        "MANITOBA" :"MB", 
        "NEW BRUNSWICK": "NB", 
        "NEWFOUNDLAND AND LABRADOR" :"NL", 
        "NORTHWEST TERRITORIES": "NT", 
        "NOVA SCOTIA": "NS", 
        "NUNAVUT": "NU", 
        "ONTARIO": "ON", 
        "PRINCE EDWARD ISLAND": "PE", 
        "QUEBEC":"QC" , 
        "SASKATCHEWAN":"SK" , 
        "YUKON": "YT" }
        if address[-1] in last_p2:
            province1= str(address[-3]+" "+address[-2]+" "+address[-1])
        elif address[-1] in last_p1:
            province1= str(address[-2]+" "+address[-1])
        else:
            province1= str(address[-1])
        if province1 in short_p:
            state= short_p.get(province1)
        else:
            state= province1
        dele_pro=(province1).split(" ")
        new_add= []
        for i in range(len(address)):
            if address[i] not in dele_pro:
                new_add.append(address[i])
        #for city:
        city= ["Falls","Payne","Grande","Bend","Vista","Post","Dorado","Smith","Springs","Bluff","Buren","Memphis","Park","Vista","Bay","York","Valley–Goose","VALLEY-GOOSE",
        "Sound", "Colborne", "Catharines", "Shores", "Thomas", "Falls–Windsor", "Grace", "City","Anthony", "John’s", "Nipissing", "House", "Louise", 
        "Hawkesbury", "Lakes","Smit", "River", "Lake", "Jaw", "Albert", "Boniface", "Factory", "Vancouver", "Rock", "George", "Rupert", "Westminster", 
        "Creek", "Hat", "Deer", "McMurray", "Prairie", "Hills","Vista","Mesa","Centro", "Cerrito","Monte", "Grove","Beach", "Habra" ,"Angeles","View", 
        "Beach","Grove", "Hueneme","Cucamonga","Bernardino","Clemente", "Diego","Fernando","Francisco","Gabriel", "Jose", "Leandro", "Obispo", "Marino", 
        "Mateo","Pedro" , "Rafael","Simeon", "Ana",  "Barbara", "Clara","Clarita", "Cruz", "Monica", "Rosa", "Valley", "Francisco","Oaks", "Creek","Covina",
        "Linda", "Creek","Morgan","Springs","Junction", "Hartford", "Haven", "Britain", "Haven","London","Saybrook", "Hartford","Haven", "Locks", "Gables",
        "Park", "Castle","Coral",  "Raton", "Glade","Land","Lauderdale","Myers","Pierce","West","Wales","Augustine","Petersburg", "Point", 
        "Valley", "Robins", "Ferry", "d’Alene", "Plaines", " Moline", "Louis","Salle", "Forest", "Vernon", "Chicago","Ridge", "Island", "Holland", 
        "Chicago", "Albany","Castle"," Harmony", "Claus", "Bend", "Haute","Lafayette","Colonies","Falls", "Rapids", "Moines", "Dodge","Pleasant", 
        "Moines", "Grove", "Lodge", "Center","Kent", "KENT" "Green", "Rouge","Charles", "Iberia","Orleans", "Martinville", "Harbor", "Kent", "Isle", "Chase", 
        "Barrington", "Bedford","Adams","Hadley",'HUBERT', 'BEAUPRÉ', 'DE-BEAUPRÉ', 'FOY', 'THÉRÈSE', 'ÎLES', 'TRACY', 'RIVIÈRES', 'NORANDA', 'NORD', 
        'PIERRE', 'ST-PIERRE', 'ST-PIERRE', 'LUC', 'ST-LUC', 'ST-LUC', 'MADELEINE','Hubert', 'Beaupré', 'de-Beaupré', 'Foy', 
        'Thérèse', 'Îles', 'Tracy', 'Rivières', 'Noranda', 'Nord', 'Pierre', 'St-Pierre', 'st-Pierre', 'Luc', 'St-Luc', 'st-Luc', 'Madeleine', 
        "Bridgewater","Springfield","Hole", "Arbor","Creek","Harbor", "Lansing", "Rapids", "Pointe",
        "Mountain","Clemens", "Pleasant","Huron", "Oak", "Ignace", "Joseph", "Marie", "Wing", "Cloud", "Paul","Centre","Paul", "Christian",
        "Gibson", "Point","Eustache", "EUSTACHE", "Girardeau", "Madrid", "Charles", "Joseph", "Genevieve","Plains", "Benton","Platte", "Vegas", "May", "Orange", 
        "Lee", "Holly", "Brunswick", "Milford","City", "Hills", "Amboy", "Village", "Cruces", "Alamos", "Consequences", "Point","Aurora", 
        "Hampton", "Hills","Falls", "Neck", "Placid", "Vernon", "Paltz", "Rochelle", "Windsor", "Marie", "James",  "John",  "Hempstead","Washington",  
        " Harbor","Brook", "Glen","Seneca","Plains","Hawk", "Head", "Bern","Mount","Forks", "Heights","Cleveland", "Liverpool","Ferry", "Vernon",
        "Philadelphia","Heights", "Reno","Pass", "Day", "Grande","Oswego","Orford","Dalles","Haven", "Southampton","Castle","Hope", "Kensington",
        "College","Greenwich", "Providence", "Kingstown","Kingstown", "Pleasant", "Fourche", "Station", "Christi", "Rio", "Pass","Paso","Worth",
        "Arthur", "Lavaca", "Angelo","Antonio", "Felipe", "Marcos" ,"Fork", "River","George", "Fork","Albans","Johnsbury","Church", "Market",
        "News","Dam", "Harbor", "Roberts","Angeles","Walla","Town","Ferry","Martinsville","Pleasant","Charleston","Claire","Lac", "Crosse", 
        "Geneva", "Glarus", "Chien", "Bend","Allis","Dells","Sleep",'FALLS', 'PAYNE', 'GRANDE', 'BEND', 'VISTA', 'POST', 'DORADO', 'SMITH', 
        'SPRINGS', 'BLUFF', 'BUREN', 'MEMPHIS', 'PARK', 'VISTA', 'BAY', 'YORK', 'SOUND', 'COLBORNE', 'CATHARINES', 'SHORES', 'THOMAS', 'FALLS–WINDSOR', 
        'GRACE', 'CITY', 'ANTHONY', 'JOHN’S', 'NIPISSING', 'HOUSE', 'LOUISE', 'HAWKESBURY', 'LAKES', 'SMIT', 'RIVER', 'LAKE', 'JAW', 'ALBERT', 
        'BONIFACE', 'FACTORY', 'VANCOUVER', 'ROCK', 'GEORGE', 'RUPERT', 'WESTMINSTER', 'CREEK', 'HAT', 'DEER', 'MCMURRAY', 'PRAIRIE', 'HILLS', 
        'VISTA', 'MESA', 'CENTRO', 'CERRITO', 'MONTE', 'GROVE', 'BEACH', 'HABRA', 'ANGELES', 'VIEW', 'BEACH', 'GROVE', 'HUENEME', 
        'CUCAMONGA', 'BERNARDINO', 'CLEMENTE', 'DIEGO', 'FERNANDO', 'FRANCISCO', 'GABRIEL', 'JOSE', 'LEANDRO', 
        'OBISPO', 'MARINO', 'MATEO', 'PEDRO', 'RAFAEL', 'SIMEON', 'ANA', 'BARBARA', 'CLARA', 'CLARITA', 'CRUZ', 'MONICA', 'ROSA', 'VALLEY', 
        'FRANCISCO', 'OAKS', 'CREEK', 'COVINA', 'LINDA', 'CREEK', 'MORGAN', 'SPRINGS', 'JUNCTION', 'HARTFORD', 'HAVEN', 'BRITAIN', 'HAVEN', 
        'LONDON', 'SAYBROOK', 'HARTFORD', 'HAVEN', 'LOCKS', 'GABLES', 'PARK', 'CASTLE', 'CORAL', 'RATON', 'GLADE', 'LAND', 'LAUDERDALE', 'MYERS', 
        'PIERCE', 'WEST', 'WALES', 'AUGUSTINE', 'PETERSBURG', 'POINT', 'VALLEY', 'ROBINS', 'FERRY', 'D’ALENE', 'PLAINES', ' MOLINE', 'LOUIS', 
        'SALLE', 'FOREST', 'VERNON', 'CHICAGO', 'RIDGE', 'ISLAND', 'HOLLAND', 'CHICAGO', 'ALBANY', 'CASTLE', ' HARMONY', 'CLAUS', 'BEND', 'HAUTE', 
        'LAFAYETTE', 'COLONIES', 'FALLS', 'RAPIDS', 'MOINES', 'DODGE', 'PLEASANT', 'MOINES', 'GROVE', 'LODGE', 'CENTER', 'GREEN', 'ROUGE', 'CHARLES', 
        'IBERIA', 'ORLEANS', 'MARTINVILLE', 'HARBOR', 'KENT', 'ISLE', 'CHASE', 'BARRINGTON', 'BEDFORD', 'ADAMS', 'HADLEY', 'BRIDGEWATER', 
        'SPRINGFIELD', 'HOLE', 'ARBOR', 'CREEK', 'HARBOR', 'LANSING', 'RAPIDS', 'POINTE', 'MOUNTAIN', 'CLEMENS', 'PLEASANT', 'HURON', 'OAK', 
        'IGNACE', 'JOSEPH', 'MARIE', 'WING', 'CLOUD', 'PAUL', 'CENTRE', 'PAUL', 'CHRISTIAN', 'GIBSON', 'POINT', 'GIRARDEAU', 'MADRID', 'CHARLES', 
        'JOSEPH', 'GENEVIEVE', 'PLAINS', 'BENTON', 'PLATTE', 'VEGAS', 'MAY', 'ORANGE', 'LEE', 'HOLLY', 'BRUNSWICK', 'MILFORD', 'CITY', 'HILLS', 
        'AMBOY', 'VILLAGE', 'CRUCES', 'ALAMOS', 'CONSEQUENCES', 'POINT', 'AURORA', 'HAMPTON', 'HILLS', 'FALLS', 'NECK', 'PLACID', 'VERNON', 
        'PALTZ', 'ROCHELLE', 'WINDSOR', 'MARIE', 'JAMES', 'JOHN', 'HEMPSTEAD', 'WASHINGTON', ' HARBOR', 'BROOK', 'GLEN', 'SENECA', 'PLAINS', 
        'HAWK', 'HEAD', 'BERN', 'MOUNT', 'FORKS', 'HEIGHTS', 'CLEVELAND', 'LIVERPOOL', 'FERRY', 'VERNON', 'PHILADELPHIA', 'HEIGHTS', 'RENO', 
        'PASS', 'DAY', 'GRANDE', 'OSWEGO', 'ORFORD', 'DALLES', 'HAVEN', 'SOUTHAMPTON', 'CASTLE', 'HOPE', 'KENSINGTON', 'COLLEGE', 'GREENWICH', 
        'PROVIDENCE', 'KINGSTOWN', 'KINGSTOWN', 'PLEASANT', 'FOURCHE', 'STATION', 'CHRISTI', 'RIO', 'PASS', 'PASO', 'WORTH', 'ARTHUR', 'LAVACA', 
        'ANGELO', 'ANTONIO', 'FELIPE', 'MARCOS', 'FORK', 'RIVER', 'GEORGE', 'FORK', 'ALBANS', 'JOHNSBURY', 'CHURCH', 'MARKET', 'NEWS', 'DAM', 'HARBOR', 
        'ROBERTS', 'ANGELES', 'WALLA', 'TOWN', 'FERRY', 'MARTINSVILLE', 'PLEASANT', 'CHARLESTON', 'CLAIRE', 'LAC', 'CROSSE', 'GENEVA', 'GLARUS', 'CHIEN', 
        'BEND', 'ALLIS', 'DELLS', 'SLEEP']
        Mid_city= ['ST', 'SAINT', 'AUX', 'VALLEY–GOOSE','GOOSE','Goose', 'ANNE','St', 'Saint', 'aux', 'Valley–Goose', 'Anne']
        if new_add[-1] in city:
            if new_add[-2] in Mid_city:
                City= str(new_add[-3]+" "+new_add[-2]+" "+new_add[-1])
            else:
                City= str(new_add[-2]+" "+new_add[-1])
        else:
            City= str(new_add[-1])
        dele_City=(City).split(" ")
        new_add1= []
        for i in range(len(new_add)):
            if new_add[i] not in dele_City:
                new_add1.append(new_add[i])
        #for address split
        strr= " ".join(new_add1)
        f_strr= (strr + " ")
        index= 0
        lst= []
        while index<len(f_strr):
            if f_strr[index]== " ":
                lst.append(index)
            index+=1
        for i in lst:
            if lst[-1]<=25:
                add_lin_1= f_strr
            elif lst[-2]<=25:
                add_lin_1= " ".join(new_add1[0:len(lst)-1])         
            elif lst[-3]<=25:
                add_lin_1= " ".join(new_add1[0:len(lst)-2])         
            elif lst[-4]<=25:
                add_lin_1= " ".join(new_add1[0:len(lst)-3])
            elif lst[-5]<=25:
                add_lin_1= " ".join(new_add1[0:len(lst)-4])
            elif lst[-6]<=25:
                add_lin_1= " ".join(new_add1[0:len(lst)-5])
            elif lst[-7]<=25:
                add_lin_1= " ".join(new_add1[0:len(lst)-6])
            elif lst[-8]<=25:
                add_lin_1= " ".join(new_add1[0:len(lst)-7])
            elif lst[-9]<=25:
                add_lin_1= " ".join(new_add1[0:len(lst)-8])
            elif lst[-10]<=25:
                add_lin_1= " ".join(new_add1[0:len(lst)-9])
            elif lst[-11]<=25:
                add_lin_1= " ".join(new_add1[0:len(lst)-10])
            elif lst[-12]<=25:
                add_lin_1= " ".join(new_add1[0:len(lst)-11])         
            elif lst[-13]<=25:
                add_lin_1= " ".join(new_add1[0:len(lst)-12])         
            elif lst[-14]<=25:
                add_lin_1= " ".join(new_add1[0:len(lst)-13])
            elif lst[-15]<=25:
                add_lin_1= " ".join(new_add1[0:len(lst)-14])
            elif lst[-16]<=25:
                add_lin_1= " ".join(new_add1[0:len(lst)-15])
            elif lst[-17]<=25:
                add_lin_1= " ".join(new_add1[0:len(lst)-16])
            elif lst[-18]<=25:
                add_lin_1= " ".join(new_add1[0:len(lst)-17])
            elif lst[-19]<=25:
                add_lin_1= " ".join(new_add1[0:len(lst)-18])    
        g=add_lin_1.split(" ")
        last_line= []
        for i in new_add1:
            if i not in g:
                last_line.append(i)
        LAST_LINE=" ".join(last_line)
        lak.address1.set(add_lin_1)
        lak.address2.set(LAST_LINE)
        lak.city.set(City)
        lak.Province.set(state)
        lak.Zip_code.set(pin)
        City= "{:<25}{}".format(City, state)
        lak.address_to_copy= (add_lin_1+"\n"+LAST_LINE+"\n"+City+"\n"+pin)


if __name__== "__main__":
    root=Tk()
    aap= Coi_Calci(root)
    root.mainloop()
    



