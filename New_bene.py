from tkinter import *
from tkinter import ttk
from tkcalendar import *
from tkcalendar import DateEntry
import pyperclip
import pandas as pd
from pandas import DataFrame as df
import holidays
from datetime import datetime
import webbrowser


class Bene_tool:
    
    def __init__(lak,root):
        lak.root= root
        lak.root.title("Beneficiary Tool")
        lak.root.geometry("1000x775")
        lak.root.resizable(False,False)
        lak.Name01= StringVar()
        lak.Releation01= StringVar()
        lak.percentage01= StringVar()
        lak.Name02= StringVar()
        lak.Releation02= StringVar()
        lak.percentage02= StringVar()
        lak.Name03= StringVar()
        lak.Releation03= StringVar()
        lak.percentage03= StringVar()
        lak.Name04= StringVar()
        lak.Releation04= StringVar()
        lak.percentage04= StringVar()
        lak.Name05= StringVar()
        lak.Releation05= StringVar()
        lak.percentage05= StringVar()
        lak.Name06= StringVar()
        lak.Releation06= StringVar()
        lak.percentage06= StringVar()
        lak.percentage2= StringVar()
        lak.relation2= StringVar()
        lak.percentage3= StringVar()
        lak.relation3= StringVar()
        lak.percentage4= StringVar()
        lak.relation4= StringVar()
        lak.Name1= StringVar()
        lak.Name2= StringVar()
        lak.Name3= StringVar()
        lak.Name4= StringVar()
        lak.Name5= StringVar()
        lak.Name6= StringVar()
        lak.Name7= StringVar()
        lak.Name8= StringVar()
        lak.Name9= StringVar()
        lak.Name10= StringVar()
        lak.Name11= StringVar()
        lak.Name12= StringVar()
        lak.detail= StringVar()
        lak.TruName1= StringVar()
        lak.TruName2= StringVar()
        lak.and_or= StringVar()
        lak.var_province= StringVar()
        lak.Var_address= StringVar()
        lak.address1= StringVar()
        lak.address2= StringVar()
        lak.city= StringVar()
        lak.Province= StringVar()
        lak.Zip_code= StringVar()
        lak.bday_date= StringVar()
        lak.bday_date10= StringVar()
        lak.bday_date20= StringVar()
        lak.Lday_date62= StringVar()
        lak.number_days= StringVar()
        lak.Main_fram()
        lak.percentage2.set(0)
        lak.percentage3.set(0)
        lak.percentage4.set(0)
    #----------------------------------------------------------------------------------------------Address Screen--------------------------------------------------------------------------------------------------------------
    def GUI(lak):
        # backgroung image
        label_bg=Label(lak.root,borderwidth=0, background="#53085C")
        label_bg.place(x=0,y=0,relwidth=1,relheight=1)
        # frame for Addreess entry box
        lak.frame= Frame(lak.root,background="#53085C")
        lak.frame.place (x=0,y=70,width=1000,height=400)
        lak.frame1= Label(lak.root,background="#53085C", text= "Enter Address:", font = ("Arial", 15, "bold"), fg= "white")
        lak.frame1.place (x=10,y=40)
        # Entrybox for Addreess
        entry =Entry(lak.frame, width=53,textvariable=lak.Var_address, borderwidth=10,foreground= "black",font = ("Arial", 25, "bold"))
        entry.place(x=15,y=0,height=50)
        #Split Button
        button1= Button(lak.frame, command=lak.Split, activebackground= "#7FB49A",activeforeground= "white", 
        text = " Split ",fg= "white",width=41, bg= "#3E0A44",font = ("Arial", 15, "bold"))
        button1.place(x=0,y=50)
        #Reset Button
        button1= Button(lak.frame, command=lak.reset, activebackground= "#7FB49A",activeforeground= "white", 
        text = " Reset ",fg= "white",width=41, bg= "#3E0A44",font = ("Arial", 15, "bold"))
        button1.place(x=502,y=50)
        #Search
        button4= Button(lak.root,activebackground= "#7FB49A",activeforeground= "white", highlightcolor="black",
        text = " BENEFICIARY TOOL ",fg= "black",command=lak.Main_fram,width=41,font = ("Arial", 15, "bold"))
        button4.place(x=0,y=0)
        button5= Button(lak.root,  activebackground= "#7FB49A",activeforeground= "white", background="#AC5CB5",
        text = " ADDRESS/CALANDER ",command= lak.GUI,fg= "black",width=41,font = ("Arial", 15, "bold"))
        button5.place(x=502,y=0)
        Copybutton= Button(lak.frame,command=lak.callback,  activebackground= "#7FB49A",activeforeground= "white", 
        text = " Google Search ",fg= "white",width=41, bg= "#3E0A44",font = ("Arial", 15, "bold"))
        Copybutton.place(x=0,y=350)
        Copybutton= Button(lak.frame,command=lak.callback_canada,  activebackground= "#7FB49A",activeforeground= "white", 
        text = " Canada Post",fg= "white",width=41, bg= "#3E0A44",font = ("Arial", 15, "bold"))
        Copybutton.place(x=502,y=350)
        #Label for Address line 1
        Line= Label(lak.frame,background="#53085C", text= "Line1 :", font = ("Arial", 12, "bold"), fg= "white")
        Line.place(x=10,y=106)
        # Entrybox for Address line 1
        Line1 =Entry(lak.frame,textvariable= lak.address1,bg="#C78BCE", width=75,foreground= "black",font = ("Arial", 14, "bold"),borderwidth=10)
        Line1.place(x=120,y=100)
        #Label for Address line 2
        Line= Label(lak.frame,background="#53085C", text= "Line2 :", font = ("Arial", 12, "bold"), fg= "white")
        Line.place(x=10,y=156)
        # Entrybox for Address line 2
        lak.Line2 =Entry(lak.frame,textvariable= lak.address2, width=75,foreground= "black",font = ("Arial", 14, "bold"),borderwidth=10,bg="#C78BCE")
        lak.Line2.place(x=120,y=150)
        #Label for City
        Zip_code= Label(lak.frame,background="#53085C", text= "City :", font = ("Arial", 12, "bold"), fg= "white")
        Zip_code.place(x=10,y=206)
        # Entrybox for City
        Zip_code =Entry(lak.frame,textvariable= lak.city , width=75,foreground= "black",font = ("Arial", 14, "bold"),borderwidth=10,bg="#C78BCE")
        Zip_code.place(x=120,y=200)
        #Label for Province
        Province= Label(lak.frame,background="#53085C", text= "Province :", font = ("Arial", 12, "bold"), fg= "white")
        Province.place(x=10,y=256)
        # Entrybox for Province
        Province =Entry(lak.frame,textvariable= lak.Province, width=75,foreground= "black",font = ("Arial", 14, "bold"),borderwidth=10,bg="#C78BCE")
        Province.place(x=120,y=250)
        #Label for Zip_code
        Zip_code= Label(lak.frame,background="#53085C", text= "Zip_code :", font = ("Arial", 12, "bold"), fg= "white")
        Zip_code.place(x=10,y=306)
        # Entrybox for Zip_code
        Zip_code =Entry(lak.frame,textvariable= lak.Zip_code , width=75,foreground= "black",font = ("Arial", 14, "bold"),borderwidth=10,bg="#C78BCE")
        Zip_code.place(x=120,y=300)
        #--------------------------------------------------------------------------------------------------------------------
        lak.frame1= Label(lak.root,background="#53085C", text= "Calander:", font = ("Arial", 15, "bold"), fg= "white")
        lak.frame1.place (x=10,y=465)
        # Label for Paid to Dat
        notes= Label(lak.root,text= " To find Lapse date => For MLIF : Select Previous Due Date \ For Ing : Select Lapse Start Date",
        justify="left",background="#53085C", font = ("Arial",9, "bold"), fg= "white")
        notes.place(x=15,y=490,width=530,height= 40)
        PTD_lbl= Label(lak.root, text= "Select Date",width=23, font = ("Arial", 12, "bold"), fg= "white", bg= "#3E0A44")
        PTD_lbl.place(x=20,y=535,height=33)
        # Entrybox for Paid to date
        PTD_lbl_entry =Entry(lak.root,  width=35,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        PTD_lbl_entry.place(x=320,y=530)
        #DATE
        cal1= DateEntry(lak.root,selectmode="day",textvariable=lak.bday_date, borderwidth=0,foreground= "black",font = ("Arial",12, "bold"), width=33)
        cal1.place(x=327,y=540)
        No_of_days= Label(lak.root, text= "Number of Days",width=23, font = ("Arial", 12, "bold"), fg= "white", bg= "#3E0A44")
        No_of_days.place(x=20,y=585, height=33)
        # Entrybox for Paid to date
        No_of_days_entry =Entry(lak.root,textvariable=lak.number_days,  width=35,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        No_of_days_entry.place(x=320,y=580)

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
        Even_date_lbl= Label(lak.root,justify="left", text= S1,bg= "white", font = ("Arial",10), fg= "white",background="#53085C")
        Even_date_lbl.place(x=750,y=520)
        Even_date_lbl= Label(lak.root,background="#53085C")
        Even_date_lbl.place(x=750,y=718,width=300,height= 30)
        Listofholi= Label(lak.root,text= "List of Holidays(Canada)",justify="right",bg= "#53085C", font = ("Arial",15, "bold"), fg= "white")
        Listofholi.place(x=722,y=470,width=300,height= 30)

        Backbutton2= Button(lak.root, command=lak.Business_Days10 , activebackground= "yellow",activeforeground= "white", 
        text = "Business Days",fg= "white",width=23, bg= "#3E0A44",font = ("Arial", 12, "bold"))
        Backbutton2.place(x=20,y=635)
        but_lbl_10entry =Entry(lak.root,textvariable=lak.bday_date10,  width=35,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        but_lbl_10entry.place(x=320,y=630)
        Backbutton3= Button(lak.root, command=lak.Business_Days20, activebackground= "yellow",activeforeground= "white", 
        text = "Regular Days",fg= "white",width=23, bg= "#3E0A44",font = ("Arial", 12, "bold"))
        Backbutton3.place(x=20,y=685)
        but_lbl_20entry =Entry(lak.root,textvariable=lak.bday_date20,  width=35,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        but_lbl_20entry.place(x=320,y=680)
        Backbutton4= Button(lak.root, command=lak.Lapse_date, activebackground= "yellow",activeforeground= "white", 
        text = "Policy Lapse Date",fg= "white",width=23, bg= "#3E0A44",font = ("Arial", 12, "bold"))
        Backbutton4.place(x=20,y=735)
        but_lbl_4entry =Entry(lak.root,textvariable=lak.Lday_date62,  width=35,borderwidth=10,foreground= "black",font = ("Arial", 12, "bold"))
        but_lbl_4entry.place(x=320,y=730)
        #Back Button
        button= Button(lak.root, command=lak.Os_reset1, activebackground= "#031b28",activeforeground= "white", 
        text = "  Reset   ",fg= "white",width=24, bg= "#3E0A44",font = ("Arial", 15, "bold"))
        button.place(x=700,y=730)
        
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
        PO_BOX= (new_add1[0])
        if PO_BOX in ["PO","Po","po","Po"]:
            add_lin_1=" ".join(new_add1[0:3])
        else:
            f_strr= (strr + " ")
            index= 0
            lst= []
            while index<len(f_strr):
                if f_strr[index]== " ":
                    lst.append(index)
                index+=1
            for i in lst:
                if lst[-1]<=30:
                    add_lin_1= f_strr
                elif lst[-2]<=30:
                    add_lin_1= " ".join(new_add1[0:len(lst)-1])         
                elif lst[-3]<=30:
                    add_lin_1= " ".join(new_add1[0:len(lst)-2])         
                elif lst[-4]<=30:
                    add_lin_1= " ".join(new_add1[0:len(lst)-3])
                elif lst[-5]<=30:
                    add_lin_1= " ".join(new_add1[0:len(lst)-4])
                elif lst[-6]<=30:
                    add_lin_1= " ".join(new_add1[0:len(lst)-5])
                elif lst[-7]<=30:
                    add_lin_1= " ".join(new_add1[0:len(lst)-6])
                elif lst[-8]<=30:
                    add_lin_1= " ".join(new_add1[0:len(lst)-7])
                elif lst[-9]<=30:
                    add_lin_1= " ".join(new_add1[0:len(lst)-8])
                elif lst[-10]<=30:
                    add_lin_1= " ".join(new_add1[0:len(lst)-9])
                elif lst[-11]<=30:
                    add_lin_1= " ".join(new_add1[0:len(lst)-10])
                elif lst[-12]<=30:
                    add_lin_1= " ".join(new_add1[0:len(lst)-11])         
                elif lst[-13]<=30:
                    add_lin_1= " ".join(new_add1[0:len(lst)-12])         
                elif lst[-14]<=30:
                    add_lin_1= " ".join(new_add1[0:len(lst)-13])
                elif lst[-15]<=30:
                    add_lin_1= " ".join(new_add1[0:len(lst)-14])
                elif lst[-16]<=30:
                    add_lin_1= " ".join(new_add1[0:len(lst)-15])
                elif lst[-17]<=30:
                    add_lin_1= " ".join(new_add1[0:len(lst)-16])
                elif lst[-18]<=30:
                    add_lin_1= " ".join(new_add1[0:len(lst)-17])
                elif lst[-19]<=30:
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
        pyperclip.copy(lak.address_to_copy)



    def Main_fram(lak):
        # backgroung image
        label_bg=Label(lak.root ,borderwidth=0, background="#268A58")
        label_bg.place(x=0,y=0,relwidth=1,relheight=1)
        #----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        #Header
        Header1= Label(lak.root,text= "Beneficiary Name",borderwidth=10,
        justify="left",bg= "#7FB49A",width=33, font = ("Arial",14, "bold"), fg= "black")
        Header1.place(x=0,y=47)
        Header1= Label(lak.root,text= "Relationship",borderwidth=10,
        justify="left",bg= "#7FB49A",width=33, font = ("Arial",14, "bold"), fg= "black")
        Header1.place(x=305,y=47)
        Header1= Label(lak.root,text= "Percentage",borderwidth=10,
        justify="left",bg= "#7FB49A",width=35, font = ("Arial",14, "bold"), fg= "black")
        Header1.place(x=625,y=47)
        
        #----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        #hidien
        ComEntry =Entry(lak.root, width=25,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#B7DECA")
        ComEntry.place(x=357,y=100)
        ComEntry =Entry(lak.root, width=25,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#B7DECA")
        ComEntry.place(x=357,y=140)
        ComEntry =Entry(lak.root, width=25,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#B7DECA")
        ComEntry.place(x=357,y=180)
        ComEntry =Entry(lak.root, width=25,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#B7DECA")
        ComEntry.place(x=357,y=220)
        ComEntry =Entry(lak.root, width=25,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#B7DECA")
        ComEntry.place(x=357,y=260)
        ComEntry =Entry(lak.root, width=25,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#B7DECA")
        ComEntry.place(x=357,y=300)
        ComEntry =Entry(lak.root, width=27,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=10,bg="#B7DECA")
        ComEntry.place(x=175,y=372)
        Nameentry =Entry(lak.root,width=10,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=10,bg="#B7DECA")
        Nameentry.place(x=505,y=560)
        # EntryBox
        Nameentry =Entry(lak.root,textvariable=lak.Name01, width=30,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#B7DECA")
        Nameentry.place(x=10,y=100)
        Que_combo= ttk.Combobox(lak.root,textvariable=lak.Releation01, width=28,font = ("Arial", 13, "bold"))
        Que_combo["values"]=("","aunts","brother","brother-in-law","brothers" ,"brothers-in-law","child" ,"children" ,"common-law husband" ,"common-law spouse"  ,"common-law wife"    ,"cousin","cousins" ,'daughter'    ,"daughter-in-law"    ,"daughters"          ,"daughters-in-law"   ,"employee"           ,"employees"          ,"ex-husband"         ,"ex-spouse"        ,"ex-wife"          ,"father"           ,"father-in-law"    ,"fiance"           ,"fiancee"          ,"foster-brother"   ,"foster-brothers"  ,"foster-child"     ,"foster-chilren"  ,"foster-daughter"  ,"foster-daughters" ,"foster-father"    ,"foster-mother"    ,"foster-parent"    ,"foster-parents"   ,"foster-sister"    ,"foster-sisters"   ,"foster-son"       ,"foster-sons"      ,"friend"         ,"friends"       ,"godchild"       ,"goddaughter"    ,"goddaughters"   ,"godfather"      ,"godmother"      ,"godparent"      ,"godparents"     ,"godson"         ,"godsons"        ,"grandchild"     ,"grandchildren"  ,"granddaughter"  ,"granddaughters" ,"grandfather"    ,"grandmother"    ,"grandparent"    ,"grandparents"        ,"grandson"            ,"grandsons"           ,"great aunt"          ,"great grandchildren" ,"great granddaughter" ,"great grandfather"   ,"great grandmother"   ,"great grandson"      ,"great nephew"        ,"great niece"         ,"great uncle"         ,"guardian"            ,"guardians"           ,"husband"             ,"mother"              ,"mother-in-law"      ,"nephew"              ,"nephews"  ,"niece"           ,"nieces"          ,"no relationship" ,"parent"          ,"parent-in-law"   ,"parents"        ,"parents-in-law"  ,"partner"         ,"partners"        ,'sibling'        ,'siblings'     ,'sister'       ,'sister-in-law'  ,'sisters'        ,'sisters-in-law' ,'son'          ,'son-in-law'     ,'sons'          ,'sons-in-law'   ,'spouse'        ,'step-brother'  ,'step-brothers'  ,'step-child'    ,'step-children' ,'step-daughter' ,'step-daughters' ,'step-father'   ,'step-mother'   ,'step-parent '  ,'step-parents','step-sisters'  ,'step-son'  ,'step-sons'   ,'trustee' ,'uncle' ,'uncles',"wife",
        "Fathers", "Mothers", "sons", "daughters", "husbands", "wifes", "brothers", "sisters", "grandfathers", "grandmothers", "grandsons", "granddaughters", "uncles", "aunts", "nephews", 
        "brother-in-law", "brothers", "brothers-in-law", "child", "children","common-law husband","common-law spouse", "common-law wife", "cousin", "cousins", "daughter-in-law", "daughters", 
        "daughters-in-law", "employee", "employees", "ex-husband","ex-spouse", "ex-wife", "father-in-law", "fiance", "fiancee", "foster-brother", "foster-brothers" ,"foster-child","Nieces")
        Que_combo.set("")
        Que_combo.place(x=362,y=106)
        PerEntry =Entry(lak.root,textvariable=lak.percentage01, width=30,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#B7DECA")
        PerEntry.place(x=650,y=100)
        #--------
        Nameentry =Entry(lak.root,textvariable=lak.Name02, width=30,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#B7DECA")
        Nameentry.place(x=10,y=140)
        Que_combo= ttk.Combobox(lak.root,textvariable=lak.Releation02, width=28,font = ("Arial", 13, "bold"))
        Que_combo["values"]=("","aunts","brother","brother-in-law","brothers" ,"brothers-in-law","child" ,"children" ,"common-law husband" ,"common-law spouse"  ,"common-law wife"    ,"cousin","cousins" ,'daughter'    ,"daughter-in-law"    ,"daughters"          ,"daughters-in-law"   ,"employee"           ,"employees"          ,"ex-husband"         ,"ex-spouse"        ,"ex-wife"          ,"father"           ,"father-in-law"    ,"fiance"           ,"fiancee"          ,"foster-brother"   ,"foster-brothers"  ,"foster-child"     ,"foster-chilren"  ,"foster-daughter"  ,"foster-daughters" ,"foster-father"    ,"foster-mother"    ,"foster-parent"    ,"foster-parents"   ,"foster-sister"    ,"foster-sisters"   ,"foster-son"       ,"foster-sons"      ,"friend"         ,"friends"       ,"godchild"       ,"goddaughter"    ,"goddaughters"   ,"godfather"      ,"godmother"      ,"godparent"      ,"godparents"     ,"godson"         ,"godsons"        ,"grandchild"     ,"grandchildren"  ,"granddaughter"  ,"granddaughters" ,"grandfather"    ,"grandmother"    ,"grandparent"    ,"grandparents"        ,"grandson"            ,"grandsons"           ,"great aunt"          ,"great grandchildren" ,"great granddaughter" ,"great grandfather"   ,"great grandmother"   ,"great grandson"      ,"great nephew"        ,"great niece"         ,"great uncle"         ,"guardian"            ,"guardians"           ,"husband"             ,"mother"              ,"mother-in-law"      ,"nephew"              ,"nephews"  ,"niece"           ,"nieces"          ,"no relationship" ,"parent"          ,"parent-in-law"   ,"parents"        ,"parents-in-law"  ,"partner"         ,"partners"        ,'sibling'        ,'siblings'     ,'sister'       ,'sister-in-law'  ,'sisters'        ,'sisters-in-law' ,'son'          ,'son-in-law'     ,'sons'          ,'sons-in-law'   ,'spouse'        ,'step-brother'  ,'step-brothers'  ,'step-child'    ,'step-children' ,'step-daughter' ,'step-daughters' ,'step-father'   ,'step-mother'   ,'step-parent '  ,'step-parents','step-sisters'  ,'step-son'  ,'step-sons'   ,'trustee' ,'uncle' ,'uncles',"wife",
        "Fathers", "Mothers", "sons", "daughters", "husbands", "wifes", "brothers", "sisters", "grandfathers", "grandmothers", "grandsons", "granddaughters", "uncles", "aunts", "nephews", 
        "brother-in-law", "brothers", "brothers-in-law", "child", "children","common-law husband","common-law spouse", "common-law wife", "cousin", "cousins", "daughter-in-law", "daughters", 
        "daughters-in-law", "employee", "employees", "ex-husband","ex-spouse", "ex-wife", "father-in-law", "fiance", "fiancee", "foster-brother", "foster-brothers" ,"foster-child","Nieces")
        Que_combo.set("")
        Que_combo.place(x=362,y=146)
        PerEntry =Entry(lak.root,textvariable=lak.percentage02, width=30,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#B7DECA")
        PerEntry.place(x=650,y=140)
        #--------
        Nameentry =Entry(lak.root,textvariable=lak.Name03, width=30,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#B7DECA")
        Nameentry.place(x=10,y=180)
        Que_combo= ttk.Combobox(lak.root,textvariable=lak.Releation03, width=28,font = ("Arial", 13, "bold"))
        Que_combo["values"]=("","aunts","brother","brother-in-law","brothers" ,"brothers-in-law","child" ,"children" ,"common-law husband" ,"common-law spouse"  ,"common-law wife"    ,"cousin","cousins" ,'daughter'    ,"daughter-in-law"    ,"daughters"          ,"daughters-in-law"   ,"employee"           ,"employees"          ,"ex-husband"         ,"ex-spouse"        ,"ex-wife"          ,"father"           ,"father-in-law"    ,"fiance"           ,"fiancee"          ,"foster-brother"   ,"foster-brothers"  ,"foster-child"     ,"foster-chilren"  ,"foster-daughter"  ,"foster-daughters" ,"foster-father"    ,"foster-mother"    ,"foster-parent"    ,"foster-parents"   ,"foster-sister"    ,"foster-sisters"   ,"foster-son"       ,"foster-sons"      ,"friend"         ,"friends"       ,"godchild"       ,"goddaughter"    ,"goddaughters"   ,"godfather"      ,"godmother"      ,"godparent"      ,"godparents"     ,"godson"         ,"godsons"        ,"grandchild"     ,"grandchildren"  ,"granddaughter"  ,"granddaughters" ,"grandfather"    ,"grandmother"    ,"grandparent"    ,"grandparents"        ,"grandson"            ,"grandsons"           ,"great aunt"          ,"great grandchildren" ,"great granddaughter" ,"great grandfather"   ,"great grandmother"   ,"great grandson"      ,"great nephew"        ,"great niece"         ,"great uncle"         ,"guardian"            ,"guardians"           ,"husband"             ,"mother"              ,"mother-in-law"      ,"nephew"              ,"nephews"  ,"niece"           ,"nieces"          ,"no relationship" ,"parent"          ,"parent-in-law"   ,"parents"        ,"parents-in-law"  ,"partner"         ,"partners"        ,'sibling'        ,'siblings'     ,'sister'       ,'sister-in-law'  ,'sisters'        ,'sisters-in-law' ,'son'          ,'son-in-law'     ,'sons'          ,'sons-in-law'   ,'spouse'        ,'step-brother'  ,'step-brothers'  ,'step-child'    ,'step-children' ,'step-daughter' ,'step-daughters' ,'step-father'   ,'step-mother'   ,'step-parent '  ,'step-parents','step-sisters'  ,'step-son'  ,'step-sons'   ,'trustee' ,'uncle' ,'uncles',"wife",
        "Fathers", "Mothers", "sons", "daughters", "husbands", "wifes", "brothers", "sisters", "grandfathers", "grandmothers", "grandsons", "granddaughters", "uncles", "aunts", "nephews", 
        "brother-in-law", "brothers", "brothers-in-law", "child", "children","common-law husband","common-law spouse", "common-law wife", "cousin", "cousins", "daughter-in-law", "daughters", 
        "daughters-in-law", "employee", "employees", "ex-husband","ex-spouse", "ex-wife", "father-in-law", "fiance", "fiancee", "foster-brother", "foster-brothers" ,"foster-child","Nieces")
        Que_combo.set("")
        Que_combo.place(x=362,y=186)
        PerEntry =Entry(lak.root,textvariable=lak.percentage03, width=30,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#B7DECA")
        PerEntry.place(x=650,y=180)
        #--------
        Nameentry =Entry(lak.root,textvariable=lak.Name04, width=30,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#B7DECA")
        Nameentry.place(x=10,y=220)
        Que_combo= ttk.Combobox(lak.root,textvariable=lak.Releation04, width=28,font = ("Arial", 13, "bold") )
        Que_combo["values"]=("","aunts","brother","brother-in-law","brothers" ,"brothers-in-law","child" ,"children" ,"common-law husband" ,"common-law spouse"  ,"common-law wife"    ,"cousin","cousins" ,'daughter'    ,"daughter-in-law"    ,"daughters"          ,"daughters-in-law"   ,"employee"           ,"employees"          ,"ex-husband"         ,"ex-spouse"        ,"ex-wife"          ,"father"           ,"father-in-law"    ,"fiance"           ,"fiancee"          ,"foster-brother"   ,"foster-brothers"  ,"foster-child"     ,"foster-chilren"  ,"foster-daughter"  ,"foster-daughters" ,"foster-father"    ,"foster-mother"    ,"foster-parent"    ,"foster-parents"   ,"foster-sister"    ,"foster-sisters"   ,"foster-son"       ,"foster-sons"      ,"friend"         ,"friends"       ,"godchild"       ,"goddaughter"    ,"goddaughters"   ,"godfather"      ,"godmother"      ,"godparent"      ,"godparents"     ,"godson"         ,"godsons"        ,"grandchild"     ,"grandchildren"  ,"granddaughter"  ,"granddaughters" ,"grandfather"    ,"grandmother"    ,"grandparent"    ,"grandparents"        ,"grandson"            ,"grandsons"           ,"great aunt"          ,"great grandchildren" ,"great granddaughter" ,"great grandfather"   ,"great grandmother"   ,"great grandson"      ,"great nephew"        ,"great niece"         ,"great uncle"         ,"guardian"            ,"guardians"           ,"husband"             ,"mother"              ,"mother-in-law"      ,"nephew"              ,"nephews"  ,"niece"           ,"nieces"          ,"no relationship" ,"parent"          ,"parent-in-law"   ,"parents"        ,"parents-in-law"  ,"partner"         ,"partners"        ,'sibling'        ,'siblings'     ,'sister'       ,'sister-in-law'  ,'sisters'        ,'sisters-in-law' ,'son'          ,'son-in-law'     ,'sons'          ,'sons-in-law'   ,'spouse'        ,'step-brother'  ,'step-brothers'  ,'step-child'    ,'step-children' ,'step-daughter' ,'step-daughters' ,'step-father'   ,'step-mother'   ,'step-parent '  ,'step-parents','step-sisters'  ,'step-son'  ,'step-sons'   ,'trustee' ,'uncle' ,'uncles',"wife",
        "Fathers", "Mothers", "sons", "daughters", "husbands", "wifes", "brothers", "sisters", "grandfathers", "grandmothers", "grandsons", "granddaughters", "uncles", "aunts", "nephews", 
        "brother-in-law", "brothers", "brothers-in-law", "child", "children","common-law husband","common-law spouse", "common-law wife", "cousin", "cousins", "daughter-in-law", "daughters", 
        "daughters-in-law", "employee", "employees", "ex-husband","ex-spouse", "ex-wife", "father-in-law", "fiance", "fiancee", "foster-brother", "foster-brothers" ,"foster-child","Nieces")
        Que_combo.set("")
        Que_combo.place(x=362,y=226)
        PerEntry =Entry(lak.root,textvariable=lak.percentage04, width=30,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#B7DECA")
        PerEntry.place(x=650,y=220)
        #--------
        Nameentry =Entry(lak.root,textvariable=lak.Name05, width=30,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#B7DECA")
        Nameentry.place(x=10,y=260)
        Que_combo= ttk.Combobox(lak.root,textvariable=lak.Releation05, width=28,font = ("Arial", 13, "bold") )
        Que_combo["values"]=("","aunts","brother","brother-in-law","brothers" ,"brothers-in-law","child" ,"children" ,"common-law husband" ,"common-law spouse"  ,"common-law wife"    ,"cousin","cousins" ,'daughter'    ,"daughter-in-law"    ,"daughters"          ,"daughters-in-law"   ,"employee"           ,"employees"          ,"ex-husband"         ,"ex-spouse"        ,"ex-wife"          ,"father"           ,"father-in-law"    ,"fiance"           ,"fiancee"          ,"foster-brother"   ,"foster-brothers"  ,"foster-child"     ,"foster-chilren"  ,"foster-daughter"  ,"foster-daughters" ,"foster-father"    ,"foster-mother"    ,"foster-parent"    ,"foster-parents"   ,"foster-sister"    ,"foster-sisters"   ,"foster-son"       ,"foster-sons"      ,"friend"         ,"friends"       ,"godchild"       ,"goddaughter"    ,"goddaughters"   ,"godfather"      ,"godmother"      ,"godparent"      ,"godparents"     ,"godson"         ,"godsons"        ,"grandchild"     ,"grandchildren"  ,"granddaughter"  ,"granddaughters" ,"grandfather"    ,"grandmother"    ,"grandparent"    ,"grandparents"        ,"grandson"            ,"grandsons"           ,"great aunt"          ,"great grandchildren" ,"great granddaughter" ,"great grandfather"   ,"great grandmother"   ,"great grandson"      ,"great nephew"        ,"great niece"         ,"great uncle"         ,"guardian"            ,"guardians"           ,"husband"             ,"mother"              ,"mother-in-law"      ,"nephew"              ,"nephews"  ,"niece"           ,"nieces"          ,"no relationship" ,"parent"          ,"parent-in-law"   ,"parents"        ,"parents-in-law"  ,"partner"         ,"partners"        ,'sibling'        ,'siblings'     ,'sister'       ,'sister-in-law'  ,'sisters'        ,'sisters-in-law' ,'son'          ,'son-in-law'     ,'sons'          ,'sons-in-law'   ,'spouse'        ,'step-brother'  ,'step-brothers'  ,'step-child'    ,'step-children' ,'step-daughter' ,'step-daughters' ,'step-father'   ,'step-mother'   ,'step-parent '  ,'step-parents','step-sisters'  ,'step-son'  ,'step-sons'   ,'trustee' ,'uncle' ,'uncles',"wife",
        "Fathers", "Mothers", "sons", "daughters", "husbands", "wifes", "brothers", "sisters", "grandfathers", "grandmothers", "grandsons", "granddaughters", "uncles", "aunts", "nephews", 
        "brother-in-law", "brothers", "brothers-in-law", "child", "children","common-law husband","common-law spouse", "common-law wife", "cousin", "cousins", "daughter-in-law", "daughters", 
        "daughters-in-law", "employee", "employees", "ex-husband","ex-spouse", "ex-wife", "father-in-law", "fiance", "fiancee", "foster-brother", "foster-brothers" ,"foster-child","Nieces")
        Que_combo.set("")
        Que_combo.place(x=362,y=266)
        PerEntry =Entry(lak.root,textvariable=lak.percentage05, width=30,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#B7DECA")
        PerEntry.place(x=650,y=260)
        #--------
        Nameentry =Entry(lak.root,textvariable=lak.Name06, width=30,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#B7DECA")
        Nameentry.place(x=10,y=300)
        Que_combo= ttk.Combobox(lak.root,textvariable=lak.Releation06, width=28,font = ("Arial", 13, "bold") )
        Que_combo["values"]=("","aunts","brother","brother-in-law","brothers" ,"brothers-in-law","child" ,"children" ,"common-law husband" ,"common-law spouse"  ,"common-law wife"    ,"cousin","cousins" ,'daughter'    ,"daughter-in-law"    ,"daughters"          ,"daughters-in-law"   ,"employee"           ,"employees"          ,"ex-husband"         ,"ex-spouse"        ,"ex-wife"          ,"father"           ,"father-in-law"    ,"fiance"           ,"fiancee"          ,"foster-brother"   ,"foster-brothers"  ,"foster-child"     ,"foster-chilren"  ,"foster-daughter"  ,"foster-daughters" ,"foster-father"    ,"foster-mother"    ,"foster-parent"    ,"foster-parents"   ,"foster-sister"    ,"foster-sisters"   ,"foster-son"       ,"foster-sons"      ,"friend"         ,"friends"       ,"godchild"       ,"goddaughter"    ,"goddaughters"   ,"godfather"      ,"godmother"      ,"godparent"      ,"godparents"     ,"godson"         ,"godsons"        ,"grandchild"     ,"grandchildren"  ,"granddaughter"  ,"granddaughters" ,"grandfather"    ,"grandmother"    ,"grandparent"    ,"grandparents"        ,"grandson"            ,"grandsons"           ,"great aunt"          ,"great grandchildren" ,"great granddaughter" ,"great grandfather"   ,"great grandmother"   ,"great grandson"      ,"great nephew"        ,"great niece"         ,"great uncle"         ,"guardian"            ,"guardians"           ,"husband"             ,"mother"              ,"mother-in-law"      ,"nephew"              ,"nephews"  ,"niece"           ,"nieces"          ,"no relationship" ,"parent"          ,"parent-in-law"   ,"parents"        ,"parents-in-law"  ,"partner"         ,"partners"        ,'sibling'        ,'siblings'     ,'sister'       ,'sister-in-law'  ,'sisters'        ,'sisters-in-law' ,'son'          ,'son-in-law'     ,'sons'          ,'sons-in-law'   ,'spouse'        ,'step-brother'  ,'step-brothers'  ,'step-child'    ,'step-children' ,'step-daughter' ,'step-daughters' ,'step-father'   ,'step-mother'   ,'step-parent '  ,'step-parents','step-sisters'  ,'step-son'  ,'step-sons'   ,'trustee' ,'uncle' ,'uncles',"wife",
        "Fathers", "Mothers", "sons", "daughters", "husbands", "wifes", "brothers", "sisters", "grandfathers", "grandmothers", "grandsons", "granddaughters", "uncles", "aunts", "nephews", 
        "brother-in-law", "brothers", "brothers-in-law", "child", "children","common-law husband","common-law spouse", "common-law wife", "cousin", "cousins", "daughter-in-law", "daughters", 
        "daughters-in-law", "employee", "employees", "ex-husband","ex-spouse", "ex-wife", "father-in-law", "fiance", "fiancee", "foster-brother", "foster-brothers" ,"foster-child","Nieces")
        Que_combo.set("")
        Que_combo.place(x=362,y=306)
        PerEntry =Entry(lak.root, textvariable=lak.percentage06,width=30,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#B7DECA")
        PerEntry.place(x=650,y=300)
        #----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        #Header
        Header= Label(lak.root,text= "For Clubbing:",
        justify="left",bg= "#268A58", font = ("Arial",15, "bold"), fg= "black")
        Header.place(x=10,y=340)
        Header1= Label(lak.root,text= " Relationship   ",borderwidth=10,
        justify="left",bg= "#7FB49A", font = ("Arial",15, "bold"), fg= "black")
        Header1.place(x=10,y=372)
        Header1= Label(lak.root,text= "Percentage  ",borderwidth=10,
        justify="left",bg= "#7FB49A", font = ("Arial",15, "bold"), fg= "black")
        Header1.place(x=505,y=372)
        
        #----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Que_combo= ttk.Combobox(lak.root,textvariable=lak.relation2, width=23,font = ("Arial", 16, "bold") )
        Que_combo["values"]=("","aunts","brother","brother-in-law","brothers" ,"brothers-in-law","child" ,"children" ,"common-law husband" ,"common-law spouse"  ,"common-law wife"    ,"cousin","cousins" ,'daughter'    ,"daughter-in-law"    ,"daughters"          ,"daughters-in-law"   ,"employee"           ,"employees"          ,"ex-husband"         ,"ex-spouse"        ,"ex-wife"          ,"father"           ,"father-in-law"    ,"fiance"           ,"fiancee"          ,"foster-brother"   ,"foster-brothers"  ,"foster-child"     ,"foster-chilren"  ,"foster-daughter"  ,"foster-daughters" ,"foster-father"    ,"foster-mother"    ,"foster-parent"    ,"foster-parents"   ,"foster-sister"    ,"foster-sisters"   ,"foster-son"       ,"foster-sons"      ,"friend"         ,"friends"       ,"godchild"       ,"goddaughter"    ,"goddaughters"   ,"godfather"      ,"godmother"      ,"godparent"      ,"godparents"     ,"godson"         ,"godsons"        ,"grandchild"     ,"grandchildren"  ,"granddaughter"  ,"granddaughters" ,"grandfather"    ,"grandmother"    ,"grandparent"    ,"grandparents"        ,"grandson"            ,"grandsons"           ,"great aunt"          ,"great grandchildren" ,"great granddaughter" ,"great grandfather"   ,"great grandmother"   ,"great grandson"      ,"great nephew"        ,"great niece"         ,"great uncle"         ,"guardian"            ,"guardians"           ,"husband"             ,"mother"              ,"mother-in-law"      ,"nephew"              ,"nephews"  ,"niece"           ,"nieces"          ,"no relationship" ,"parent"          ,"parent-in-law"   ,"parents"        ,"parents-in-law"  ,"partner"         ,"partners"        ,'sibling'        ,'siblings'     ,'sister'       ,'sister-in-law'  ,'sisters'        ,'sisters-in-law' ,'son'          ,'son-in-law'     ,'sons'          ,'sons-in-law'   ,'spouse'        ,'step-brother'  ,'step-brothers'  ,'step-child'    ,'step-children' ,'step-daughter' ,'step-daughters' ,'step-father'   ,'step-mother'   ,'step-parent '  ,'step-parents','step-sisters'  ,'step-son'  ,'step-sons'   ,'trustee' ,'uncle' ,'uncles',"wife",
        "Fathers", "Mothers", "sons", "daughters", "husbands", "wifes", "brothers", "sisters", "grandfathers", "grandmothers", "grandsons", "granddaughters", "uncles", "aunts", "nephews", 
        "brother-in-law", "brothers", "brothers-in-law", "child", "children","common-law husband","common-law spouse", "common-law wife", "cousin", "cousins", "daughter-in-law", "daughters", 
        "daughters-in-law", "employee", "employees", "ex-husband","ex-spouse", "ex-wife", "father-in-law", "fiance", "fiancee", "foster-brother", "foster-brothers" ,"foster-child","Nieces")
        Que_combo.set("")
        Que_combo.place(x=182,y=380)
        PerEntry =Entry(lak.root,textvariable=lak.percentage2, width=30,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=10,bg="#B7DECA")
        PerEntry.place(x=640,y=372)
        #----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Nameentry =Entry(lak.root,textvariable=lak.Name1, width=28,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#B7DECA")
        Nameentry.place(x=10,y=425)
        Nameentry =Entry(lak.root,textvariable=lak.Name2, width=28,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#B7DECA")
        Nameentry.place(x=340,y=425)
        Nameentry =Entry(lak.root,textvariable=lak.Name3, width=28,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#B7DECA")
        Nameentry.place(x=670,y=425)
        Nameentry =Entry(lak.root,textvariable=lak.Name4, width=28,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#B7DECA")
        Nameentry.place(x=10,y=465)
        Nameentry =Entry(lak.root,textvariable=lak.Name5, width=28,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#B7DECA")
        Nameentry.place(x=340,y=465)
        Nameentry =Entry(lak.root,textvariable=lak.Name6, width=28,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#B7DECA")
        Nameentry.place(x=670,y=465)
        Nameentry =Entry(lak.root,textvariable=lak.Name7, width=28,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#B7DECA")
        Nameentry.place(x=10,y=505)
        Nameentry =Entry(lak.root,textvariable=lak.Name8, width=28,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#B7DECA")
        Nameentry.place(x=340,y=505)
        Nameentry =Entry(lak.root,textvariable=lak.Name9, width=28,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#B7DECA")
        Nameentry.place(x=670,y=505)
        #----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        #----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Header1= Label(lak.root,text= " Trustee(s)  ",borderwidth=10,
        justify="left",bg= "#7FB49A", font = ("Arial",15, "bold"), fg= "black")
        Header1.place(x=10,y=560)
        Nameentry =Entry(lak.root,textvariable=lak.TruName1, width=30,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=10,bg="#B7DECA")
        Nameentry.place(x=150,y=560)
        Que_combo= ttk.Combobox(lak.root,textvariable=lak.and_or, width=8,state= "readonly",font = ("Arial",15, "bold") )
        Que_combo["values"]=("","and","or")
        Que_combo.set("")
        Que_combo.place(x=515,y=568)
        Nameentry =Entry(lak.root,textvariable=lak.TruName2, width=30,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=10,bg="#B7DECA")
        Nameentry.place(x=640,y=560)
        #----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Header= Label(lak.root,text= "Beneficiary Detail",
        justify="left",bg= "#268A58", font = ("Arial",20, "bold"), fg= "black")
        Header.place(x=370,y=615)
        DetailEntry =Entry(lak.root,textvariable=lak.detail, width=107,foreground= "black",font = ("Arial",12, "bold"),borderwidth=10,bg="#B7DECA")
        DetailEntry.place(x=5,y=650, height=65)
        #----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        # button
        button= Button(lak.root,command= lak.benefunction,bg= "#7FB49A", activebackground= "#031b28",activeforeground= "white",
        text = "Primary Beneficiary ",fg= "black",width=27,font = ("Arial", 15, "bold"))
        button.place(x=0,y=720)
        button= Button(lak.root,command= lak.Secondry_Bene,bg= "#7FB49A", activebackground= "#031b28",activeforeground= "white",
        text = "Secondry Beneficiary ",fg= "black",width=27,font = ("Arial", 15, "bold"))
        button.place(x=335,y=720)
        button5= Button(lak.root,bg= "#7FB49A",  activebackground= "#7FB49A",activeforeground= "white", 
        text = " Reset ",command= lak.rest,fg= "black",width=27,font = ("Arial", 15, "bold"))
        button5.place(x=670,y=720)
        button4= Button(lak.root,activebackground= "#7FB49A",activeforeground= "white", bg="#7FB49A",
        text = " BENEFICIARY TOOL ",fg= "black",command=lak.Main_fram,width=41,font = ("Arial", 15, "bold"))
        button4.place(x=0,y=0)
        button5= Button(lak.root,  activebackground= "#7FB49A",activeforeground= "white", 
        text = " ADDRESS/CALANDER ",command= lak.GUI,fg= "black",width=41,font = ("Arial", 15, "bold"))
        button5.place(x=502,y=0)
    #----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    #Primary Bene'
    def benefunction(lak):
        Namelist= [lak.Name01.get(),lak.Name02.get(),lak.Name03.get(),lak.Name04.get(),lak.Name05.get(),lak.Name06.get()]
        Relationlist=[lak.Releation01.get(),lak.Releation02.get(),lak.Releation03.get(),lak.Releation04.get(),lak.Releation05.get(),lak.Releation06.get()]
        perlist=[lak.percentage01.get(),lak.percentage02.get(),lak.percentage03.get(),lak.percentage04.get(),lak.percentage05.get(),lak.percentage06.get()]
        Fnamelist= []
        FRelationlist=[]
        Fperlist=[]
        for i in range(6):
            if Namelist[i]!="":
                Fnamelist.append(Namelist[i])
        for i in range(len(Fnamelist)):
            FRelationlist.append(Relationlist[i])
            Fperlist.append(perlist[i])

        Syntax= []
        for i in range(len(Fnamelist)):
            
            N= Fnamelist[i]
            N=N.title()
            R= FRelationlist[i]
            P= Fperlist[i]
            if N != "":
                if N!="Estate":
                    if P !="":
                        P= eval(Fperlist[i])
                        if P<=100:
                            final= "({})-({}) {}%".format(N, R, P)
                        else:
                            final= "({})-({}) ${}".format(N, R, P)
                        Syntax.append(final)
                    elif P =="":
                        final= "({})-({})".format(N, R)
                        Syntax.append(final)
                elif N=="Estate":
                    if P !="":
                        P= eval(Fperlist[i])
                        if P<=100:
                            final= "{} {} {}%".format(N, R, P)
                        else:
                            final= "{} {} ${}".format(N, R, P)
                        Syntax.append(final)
                    elif P =="":
                        final= "{} {}".format(N, R)
                        Syntax.append(final)

        Syntaxs=" and ".join(Syntax)
        TN1=lak.TruName1.get()
        TN1=TN1.title()
        TNO=lak.and_or.get()
        TN2=lak.TruName2.get()
        TN2=TN2.title()
        if lak.Name1.get()=="": 
            if TN1!="":
                if TNO=="":
                    Syntaxs="{} *** Trustee - {} ***".format(Syntaxs,TN1)
                else:
                    Syntaxs="{} *** Trustee - {} {} {} ***".format(Syntaxs,TN1,TNO,TN2 )
            else:
                if P =="":
                    Syntaxs=(Syntaxs)
                else: 
                    Syntaxs=(Syntaxs+".")   

        #Clubbing
        Relation= lak.relation2.get()
        Per=eval(lak.percentage2.get())
        Name1= lak.Name1.get()
        Name1=Name1.title()
        Name2= lak.Name2.get()
        Name2=Name2.title()
        Name3= lak.Name3.get()
        Name3=Name3.title()
        Name4= lak.Name4.get()
        Name4=Name4.title()
        Name5= lak.Name5.get()
        Name5=Name5.title()
        Name6= lak.Name6.get()
        Name6=Name6.title()
        Name7= lak.Name7.get()
        Name7=Name7.title()
        Name8= lak.Name8.get()
        Name8=Name8.title()
        Name9= lak.Name9.get()
        Name9=Name9.title()


        Namelst=[Name1, Name2,Name3, Name4,Name5,Name6,Name7,Name8,Name9]
        Fname=[]
        
        for i in range(len(Namelst)):
            if Namelst[i]!="":
                Fname.append(Namelst[i])
        Fper= Per*len(Fname)
        ClubSyntax=[]
        for i in range(len(Fname)):
            N= Fname[i]
            if N=="":
                pass
            elif N != "":
                clubname= "{}".format(N)
                ClubSyntax.append(clubname)
        ClubSyntaxs= " and ".join(ClubSyntax)
        if Fper==0:
            final= "({})-({})".format(ClubSyntaxs, Relation)
        else:    
            if Fper <= 100:
                final= "({})-({}) {}%".format(ClubSyntaxs, Relation, Fper)
            else:
                final= "({})-({}) ${}".format(ClubSyntaxs, Relation, Fper)
        
        if Syntaxs=="":
            FinalSyntax=final
        else:
            FinalSyntax=Syntaxs+" and "+final
        TN1=lak.TruName1.get()
        TN1=TN1.title()
        TNO=lak.and_or.get()
        TN2=lak.TruName2.get()
        TN2=TN2.title()
        if lak.Name7.get()=="": 
            if TN1!="":
                if TNO=="":
                    FinalSyntax="{} *** Trustee - {} ***".format(FinalSyntax,TN1)
                else:
                    FinalSyntax="{} *** Trustee - {} {} {} ***".format(FinalSyntax,TN1,TNO,TN2 )
            elif Fper==0:
                FinalSyntax=(FinalSyntax)
            else:
                FinalSyntax=(FinalSyntax+".")
        
        if ClubSyntax==[]:
            lak.detail.set(Syntaxs)
            pyperclip.copy(lak.detail.get())
        elif ClubSyntax != []:
            lak.detail.set(FinalSyntax)
            pyperclip.copy(lak.detail.get())
        else:
            lak.detail.set("Please Check The Entries")

    #-------------------------------------------Secondry Benef--------------------------------------------------------------------------------------------------------------------------------------------------
    def Secondry_Bene(lak):
        Namelist= [lak.Name01.get(),lak.Name02.get(),lak.Name03.get(),lak.Name04.get(),lak.Name05.get(),lak.Name06.get()]
        Relationlist=[lak.Releation01.get(),lak.Releation02.get(),lak.Releation03.get(),lak.Releation04.get(),lak.Releation05.get(),lak.Releation06.get()]
        perlist=[lak.percentage01.get(),lak.percentage02.get(),lak.percentage03.get(),lak.percentage04.get(),lak.percentage05.get(),lak.percentage06.get()]
        Fnamelist= []
        FRelationlist=[]
        Fperlist=[]
        for i in range(6):
            if Namelist[i]!="":
                Fnamelist.append(Namelist[i])
        for i in range(len(Fnamelist)):
            FRelationlist.append(Relationlist[i])
            Fperlist.append(perlist[i])

        Syntax= []
        for i in range(len(Fnamelist)):
            
            N= Fnamelist[i]
            N=N.title()
            R= FRelationlist[i]
            P= Fperlist[i]
            if N != "":
                if N!="Estate":
                    if P !="":
                        P= eval(Fperlist[i])
                        if P<=100:
                            final= "{}-{} {}%".format(N, R, P)
                        else:
                            final= "{}-{} ${}".format(N, R, P)
                        Syntax.append(final)
                    elif P =="":
                        final= "{}-{}".format(N, R)
                        Syntax.append(final)
                elif N=="Estate":
                    if P !="":
                        P= eval(Fperlist[i])
                        if P<=100:
                            final= "{} {} {}%".format(N, R, P)
                        else:
                            final= "{} {} ${}".format(N, R, P)
                        Syntax.append(final)
                    elif P =="":
                        final= "{} {}".format(N, R)
                        Syntax.append(final)

        Syntaxs=" & ".join(Syntax)
        TN1=lak.TruName1.get()
        TN1=TN1.title()
        TNO=lak.and_or.get()
        TN2=lak.TruName2.get()
        TN2=TN2.title()
        if lak.Name1.get()=="": 
            if TN1!="":
                if TNO=="":
                    Syntaxs="{} *** Trustee - {} ***".format(Syntaxs,TN1)
                else:
                    Syntaxs="{} *** Trustee - {} {} {} ***".format(Syntaxs,TN1,TNO,TN2 )
            else:
                if P =="":
                    Syntaxs=(Syntaxs)
                else: 
                    Syntaxs=(Syntaxs+".")  

        #Clubbing
        Relation= lak.relation2.get()
        Per=eval(lak.percentage2.get())
        Name1= lak.Name1.get()
        Name1=Name1.title()
        Name2= lak.Name2.get()
        Name2=Name2.title()
        Name3= lak.Name3.get()
        Name3=Name3.title()
        Name4= lak.Name4.get()
        Name4=Name4.title()
        Name5= lak.Name5.get()
        Name5=Name5.title()
        Name6= lak.Name6.get()
        Name6=Name6.title()
        Name7= lak.Name7.get()
        Name7=Name7.title()
        Name8= lak.Name8.get()
        Name8=Name8.title()
        Name9= lak.Name9.get()
        Name9=Name9.title()


        Namelst=[Name1, Name2,Name3, Name4,Name5,Name6,Name7,Name8,Name9]
        Fname=[]
        
        for i in range(len(Namelst)):
            if Namelst[i]!="":
                Fname.append(Namelst[i])
        Fper= Per*len(Fname)
        ClubSyntax=[]
        for i in range(len(Fname)):
            N= Fname[i]
            if N=="":
                pass
            elif N != "":
                clubname= "{}".format(N)
                ClubSyntax.append(clubname)
        ClubSyntaxs= " & ".join(ClubSyntax)
        if Fper==0:
            final= "{}-{}".format(ClubSyntaxs, Relation)
        else:    
            if Fper <= 100:
                final= "{}-{} {}%".format(ClubSyntaxs, Relation, Fper)
            else:
                final= "{}-{} ${}".format(ClubSyntaxs, Relation, Fper)
        
        if Syntaxs=="":
            FinalSyntax=final
        else:
            FinalSyntax=Syntaxs+" & "+final
        TN1=lak.TruName1.get()
        TN1=TN1.title()
        TNO=lak.and_or.get()
        TN2=lak.TruName2.get()
        TN2=TN2.title()
        if lak.Name7.get()=="": 
            if TN1!="":
                if TNO=="":
                    FinalSyntax="{} *** Trustee - {} ***".format(FinalSyntax,TN1)
                else:
                    FinalSyntax="{} *** Trustee - {} {} {} ***".format(FinalSyntax,TN1,TNO,TN2 )
            elif Fper==0:
                FinalSyntax=(FinalSyntax)
            else:
                FinalSyntax=(FinalSyntax+".")
        
        if ClubSyntax==[]:
            lak.detail.set(Syntaxs)
            pyperclip.copy(lak.detail.get())
        elif ClubSyntax != []:
            lak.detail.set(FinalSyntax)
            pyperclip.copy(lak.detail.get())
        else:
            lak.detail.set("Please Check The Entries")

    #----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    def rest(lak):
        lak.Name01.set("")
        lak.Releation01.set("")
        lak.percentage01.set("")
        lak.Name02.set("")
        lak.Releation02.set("")
        lak.percentage02.set("")
        lak.Name03.set("")
        lak.Releation03.set("")
        lak.percentage03.set("")
        lak.Name04.set("")
        lak.Releation04.set("")
        lak.percentage04.set("")
        lak.Name05.set("")
        lak.Releation05.set("")
        lak.percentage05.set("")
        lak.Name06.set("")
        lak.Releation06.set("")
        lak.percentage06.set("")
        lak.relation2.set("")
        lak.Name1.set("")
        lak.Name2.set("")
        lak.Name3.set("")
        lak.Name4.set("")
        lak.Name5.set("")
        lak.Name6.set("")
        lak.Name7.set("")
        lak.Name8.set("")
        lak.Name9.set("")
        lak.detail.set("")
        lak.percentage2.set(0)
        lak.TruName1.set("")
        lak.TruName2.set("")
        lak.and_or.set("")


if __name__== "__main__":
    root=Tk()
    aap= Bene_tool(root)
    root.mainloop()