from tkinter import *
# from tkinter import ttk
# from PIL import Image,ImageTk
import pyperclip
import webbrowser

class Coi_Calci:
    
    def __init__(lak,root):
        lak.root= root
        lak.root.title("All In One Calulator")
        lak.root.geometry("500x450")
        lak.root.resizable(False,False)
        lak.var_province= StringVar()
        lak.Var_address= StringVar()
        lak.address1= StringVar()
        lak.address2= StringVar()
        lak.city= StringVar()
        lak.Province= StringVar()
        lak.Zip_code= StringVar()
        lak.GUI()
#--------------------------------------------------------------Address Screen--------------------------------------------------------------------------------------
    def GUI(lak):
        # backgroung image
        # lak.bg=ImageTk.PhotoImage(file="E:\\Sunlife\\slbg.jpg")
        # label_bg=Label(lak.root, image=lak.bg,borderwidth=0)
        # label_bg.place(x=0,y=0,relwidth=1,relheight=1)
        # frame for Addreess entry box
        lak.frame= Frame(lak.root, bg= "yellow")
        lak.frame.place (x=0,y=80,width=600,height=400)
        lak.frame1= Label(lak.root,  bg= "yellow",text= "Enter Address:", font = ("Arial", 12, "bold"), fg= "black")
        lak.frame1.place (x=10,y=50)
        # Entrybox for Addreess
        entry =Entry(lak.frame, width=41,textvariable=lak.Var_address, borderwidth=10,foreground= "black",font = ("Arial", 15, "bold"))
        entry.place(x=15,y=1,height=50)
        #Split Button
        button1= Button(lak.frame, command=lak.Split, activebackground= "yellow",activeforeground= "white", 
        text = " Split ",fg= "black",width=15, bg= "white",font = ("Arial", 12, "bold"))
        button1.place(x=13,y=50)
        #Reset Button
        button1= Button(lak.frame, command=lak.reset, activebackground= "yellow",activeforeground= "white", 
        text = " Reset ",fg= "black",width=15, bg= "white",font = ("Arial", 12, "bold"))
        button1.place(x=320,y=50)
        #Reset Button
        button1= Button(lak.frame, command=lambda:pyperclip.copy(lak.address_to_copy), activebackground= "yellow",activeforeground= "white", 
        text = " Copy ",fg= "black",width=15, bg= "white",font = ("Arial", 12, "bold"))
        button1.place(x=172,y=50)
        #Search
        Copybutton= Button(lak.root,command=lak.callback,  activebackground= "yellow",activeforeground= "white", 
        text = " Google Search ",fg= "black",width=25, bg= "white",font = ("Arial", 12, "bold"))
        Copybutton.place(x=0,y=0)
        Copybutton= Button(lak.root,command=lak.callback_canada,  activebackground= "yellow",activeforeground= "white", 
        text = " Canada Post",fg= "black",width=25, bg= "white",font = ("Arial", 12, "bold"))
        Copybutton.place(x=260,y=0)
        #Label for Address line 1
        Line= Label(lak.frame, text= "Line1 :", font = ("Arial", 12, "bold"), fg= "black")
        Line.place(x=10,y=100)
        # Entrybox for Address line 1
        Line1 =Entry(lak.frame,textvariable= lak.address1, width=30,foreground= "black",font = ("Arial", 14, "bold"),borderwidth=10)
        Line1.place(x=120,y=100)
        #Label for Address line 2
        Line= Label(lak.frame, text= "Line2 :", font = ("Arial", 12, "bold"), fg= "black")
        Line.place(x=10,y=150)
        # Entrybox for Address line 2
        lak.Line2 =Entry(lak.frame,textvariable= lak.address2, width=30,foreground= "black",font = ("Arial", 14, "bold"),borderwidth=10)
        lak.Line2.place(x=120,y=150)
        #Label for City
        Zip_code= Label(lak.frame, text= "City :", font = ("Arial", 12, "bold"), fg= "black")
        Zip_code.place(x=10,y=200)
        # Entrybox for City
        Zip_code =Entry(lak.frame,textvariable= lak.city , width=30,foreground= "black",font = ("Arial", 14, "bold"),borderwidth=10)
        Zip_code.place(x=120,y=200)
        #Label for Province
        Province= Label(lak.frame, text= "Province :", font = ("Arial", 12, "bold"), fg= "black")
        Province.place(x=10,y=250)
        # Entrybox for Province
        Province =Entry(lak.frame,textvariable= lak.Province, width=30,foreground= "black",font = ("Arial", 14, "bold"),borderwidth=10)
        Province.place(x=120,y=250)
        #Label for Zip_code
        Zip_code= Label(lak.frame, text= "Zip_code :", font = ("Arial", 12, "bold"), fg= "black")
        Zip_code.place(x=10,y=300)
        # Entrybox for Zip_code
        Zip_code =Entry(lak.frame,textvariable= lak.Zip_code , width=30,foreground= "black",font = ("Arial", 14, "bold"),borderwidth=10)
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
    



