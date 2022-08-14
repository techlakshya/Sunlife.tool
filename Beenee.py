from email.header import Header
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from PIL import Image,ImageTk
import pyperclip


class Bene_tool:
    
    def __init__(lak,root):
        lak.root= root
        lak.root.title("Beneficiary Tool")
        lak.root.geometry("1000x900")
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
        lak.Name1= StringVar()
        lak.Name2= StringVar()
        lak.Name3= StringVar()
        lak.Name4= StringVar()
        lak.Name5= StringVar()
        lak.Name6= StringVar()
        lak.detail= StringVar()
        lak.Main_fram()
        lak.percentage2.set(0)

    def Main_fram(lak):
        # backgroung image
        lak.bg=ImageTk.PhotoImage(file="E:\\Sunlife\\AdobeStock_107715529.jpg")
        label_bg=Label(lak.root, image=lak.bg,borderwidth=0)
        label_bg.place(x=0,y=0,relwidth=1,relheight=1)
        #----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        #Header
        Header= Label(lak.root,text= "BENEFICIARY TOOL",
        justify="left",bg= "#015daf", font = ("Arial",20, "bold"), fg= "white")
        Header.place(x=380,y=15)
        Header1= Label(lak.root,text= "Beneficiary Name",borderwidth=10,
        justify="left",bg= "yellow",width=30, font = ("Arial",15, "bold"), fg= "black")
        Header1.place(x=0,y=60)
        Header1= Label(lak.root,text= "Relationship",borderwidth=10,
        justify="left",bg= "yellow",width=30, font = ("Arial",15, "bold"), fg= "black")
        Header1.place(x=320,y=60)
        Header1= Label(lak.root,text= "Percentage",borderwidth=10,
        justify="left",bg= "yellow",width=30, font = ("Arial",15, "bold"), fg= "black")
        Header1.place(x=640,y=60)
        #----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        #hidien
        ComEntry =Entry(lak.root, width=25,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#fee5af")
        ComEntry.place(x=357,y=120)
        ComEntry =Entry(lak.root, width=25,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#fee5af")
        ComEntry.place(x=357,y=180)
        ComEntry =Entry(lak.root, width=25,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#fee5af")
        ComEntry.place(x=357,y=240)
        ComEntry =Entry(lak.root, width=25,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#fee5af")
        ComEntry.place(x=357,y=300)
        ComEntry =Entry(lak.root, width=25,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#fee5af")
        ComEntry.place(x=357,y=360)
        ComEntry =Entry(lak.root, width=25,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#fee5af")
        ComEntry.place(x=357,y=420)
        ComEntry =Entry(lak.root, width=27,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=10,bg="#fee5af")
        ComEntry.place(x=160,y=520)
        # EntryBox
        Nameentry =Entry(lak.root,textvariable=lak.Name01, width=30,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#fee5af")
        Nameentry.place(x=10,y=120)
        Que_combo= ttk.Combobox(lak.root,textvariable=lak.Releation01, width=28,font = ("Arial", 13, "bold"), state= "readonly")
        Que_combo["values"]=(" ","Father", "Mother", "Son", "Daughter", "Husband", "Wife", "Brother", "Sister", "Grandfather", "Grandmother", "Grandson", "Granddaughter", "Uncle", "Aunt", "Nephew", "Niece")
        Que_combo.set(" ")
        Que_combo.place(x=362,y=126)
        PerEntry =Entry(lak.root,textvariable=lak.percentage01, width=30,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#fee5af")
        PerEntry.place(x=650,y=120)
        #--------
        Nameentry =Entry(lak.root,textvariable=lak.Name02, width=30,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#fee5af")
        Nameentry.place(x=10,y=180)
        Que_combo= ttk.Combobox(lak.root,textvariable=lak.Releation02, width=28,font = ("Arial", 13, "bold"), state= "readonly")
        Que_combo["values"]=(" ","Father", "Mother", "Son", "Daughter", "Husband", "Wife", "Brother", "Sister", "Grandfather", "Grandmother", "Grandson", "Granddaughter", "Uncle", "Aunt", "Nephew", "Niece")
        Que_combo.set(" ")
        Que_combo.place(x=362,y=186)
        PerEntry =Entry(lak.root,textvariable=lak.percentage02, width=30,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#fee5af")
        PerEntry.place(x=650,y=180)
        #--------
        Nameentry =Entry(lak.root,textvariable=lak.Name03, width=30,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#fee5af")
        Nameentry.place(x=10,y=240)
        Que_combo= ttk.Combobox(lak.root,textvariable=lak.Releation03, width=28,font = ("Arial", 13, "bold"), state= "readonly")
        Que_combo["values"]=(" ","Father", "Mother", "Son", "Daughter", "Husband", "Wife", "Brother", "Sister", "Grandfather", "Grandmother", "Grandson", "Granddaughter", "Uncle", "Aunt", "Nephew", "Niece")
        Que_combo.set(" ")
        Que_combo.place(x=362,y=246)
        PerEntry =Entry(lak.root,textvariable=lak.percentage03, width=30,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#fee5af")
        PerEntry.place(x=650,y=240)
        #--------
        Nameentry =Entry(lak.root,textvariable=lak.Name04, width=30,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#fee5af")
        Nameentry.place(x=10,y=300)
        Que_combo= ttk.Combobox(lak.root,textvariable=lak.Releation04, width=28,font = ("Arial", 13, "bold"), state= "readonly")
        Que_combo["values"]=(" ","Father", "Mother", "Son", "Daughter", "Husband", "Wife", "Brother", "Sister", "Grandfather", "Grandmother", "Grandson", "Granddaughter", "Uncle", "Aunt", "Nephew", "Niece")
        Que_combo.set(" ")
        Que_combo.place(x=362,y=306)
        PerEntry =Entry(lak.root,textvariable=lak.percentage04, width=30,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#fee5af")
        PerEntry.place(x=650,y=300)
        #--------
        Nameentry =Entry(lak.root,textvariable=lak.Name05, width=30,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#fee5af")
        Nameentry.place(x=10,y=360)
        Que_combo= ttk.Combobox(lak.root,textvariable=lak.Releation05, width=28,font = ("Arial", 13, "bold"), state= "readonly")
        Que_combo["values"]=(" ","Father", "Mother", "Son", "Daughter", "Husband", "Wife", "Brother", "Sister", "Grandfather", "Grandmother", "Grandson", "Granddaughter", "Uncle", "Aunt", "Nephew", "Niece")
        Que_combo.set(" ")
        Que_combo.place(x=362,y=366)
        PerEntry =Entry(lak.root,textvariable=lak.percentage05, width=30,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#fee5af")
        PerEntry.place(x=650,y=360)
        #--------
        Nameentry =Entry(lak.root,textvariable=lak.Name06, width=30,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#fee5af")
        Nameentry.place(x=10,y=420)
        Que_combo= ttk.Combobox(lak.root,textvariable=lak.Releation06, width=28,font = ("Arial", 13, "bold"), state= "readonly")
        Que_combo["values"]=(" ","Father", "Mother", "Son", "Daughter", "Husband", "Wife", "Brother", "Sister", "Grandfather", "Grandmother", "Grandson", "Granddaughter", "Uncle", "Aunt", "Nephew", "Niece")
        Que_combo.set(" ")
        Que_combo.place(x=362,y=426)
        PerEntry =Entry(lak.root, textvariable=lak.percentage06,width=30,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#fee5af")
        PerEntry.place(x=650,y=420)
        #----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        #Header
        Header= Label(lak.root,text= "For Clubbing:",
        justify="left",bg= "#015daf", font = ("Arial",15, "bold"), fg= "yellow")
        Header.place(x=10,y=470)
        Header1= Label(lak.root,text= "Relationship",borderwidth=10,
        justify="left",bg= "yellow", font = ("Arial",15, "bold"), fg= "black")
        Header1.place(x=10,y=520)
        Header1= Label(lak.root,text= "Percentage  ",borderwidth=10,
        justify="left",bg= "yellow", font = ("Arial",15, "bold"), fg= "black")
        Header1.place(x=491,y=520)
        #----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Que_combo= ttk.Combobox(lak.root,textvariable=lak.relation2, width=23,font = ("Arial", 16, "bold"), state= "readonly")
        Que_combo["values"]=(" ","Father", "Mother", "Son", "Daughter", "Husband", "Wife", "Brother", "Sister", "Grandfather", "Grandmother", "Grandson", "Granddaughter", "Uncle", "Aunt", "Nephew", "Niece")
        Que_combo.set(" ")
        Que_combo.place(x=170,y=528)
        PerEntry =Entry(lak.root,textvariable=lak.percentage2, width=30,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=10,bg="#fee5af")
        PerEntry.place(x=640,y=520)
        #----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Nameentry =Entry(lak.root,textvariable=lak.Name1, width=43,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#fee5af")
        Nameentry.place(x=10,y=580)
        Nameentry =Entry(lak.root,textvariable=lak.Name2, width=43,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#fee5af")
        Nameentry.place(x=505,y=580)
        Nameentry =Entry(lak.root,textvariable=lak.Name3, width=43,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#fee5af")
        Nameentry.place(x=10,y=620)
        Nameentry =Entry(lak.root,textvariable=lak.Name4, width=43,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#fee5af")
        Nameentry.place(x=505,y=620)
        Nameentry =Entry(lak.root,textvariable=lak.Name5, width=43,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#fee5af")
        Nameentry.place(x=10,y=660)
        Nameentry =Entry(lak.root,textvariable=lak.Name6, width=43,foreground= "black",font = ("Arial", 15, "bold"),borderwidth=5,bg="#fee5af")
        Nameentry.place(x=505,y=660)
        #----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Header= Label(lak.root,text= "Beneficiary Detail:",
        justify="left",bg= "#015daf", font = ("Arial",16, "bold"), fg= "yellow")
        Header.place(x=10,y=705)
        DetailEntry =Entry(lak.root,textvariable=lak.detail, width=87,foreground= "black",font = ("Arial",15, "bold"),borderwidth=10,bg="#fee5af")
        DetailEntry.place(x=10,y=750)
        #----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        # button
        button= Button(lak.root,command= lak.benefunction, activebackground= "#031b28",activeforeground= "white", 
        text = " Enter ",fg= "black",width=27,font = ("Arial", 15, "bold"))
        button.place(x=10,y=820)
        button4= Button(lak.root,activebackground= "yellow",activeforeground= "white", 
        text = " Copy ",fg= "black",command=lambda:pyperclip.copy(lak.detail.get()),width=27,font = ("Arial", 15, "bold"))
        button4.place(x=345,y=820)
        button5= Button(lak.root,  activebackground= "yellow",activeforeground= "white", 
        text = " Reset ",command= lak.rest,fg= "black",width=25,font = ("Arial", 15, "bold"))
        button5.place(x=680,y=820)
        #----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        #function'
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
            R= FRelationlist[i]
            P= Fperlist[i]
            if N[i] != "":
                final= "({})-({}) {}% ".format(N, R, P)
                Syntax.append(final)
        Syntaxs=" and ".join(Syntax)

        #Clubbing
        Relation= lak.relation2.get()
        Per=float(lak.percentage2.get())
        Name1= lak.Name1.get()
        Name2= lak.Name2.get()
        Name3= lak.Name3.get()
        Name4= lak.Name4.get()
        Name5= lak.Name5.get()
        Name6= lak.Name6.get()
        Namelst=[Name1, Name2,Name3, Name4, Name5, Name6]
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
        final= "({})-({}) {}%. ".format(ClubSyntaxs, Relation, Fper)
        FinalSyntax=Syntaxs+"and"+final
        if ClubSyntax==[]:
            lak.detail.set(Syntaxs)
        elif Syntax==[]:
            lak.detail.set(final)
        else:
            lak.detail.set(FinalSyntax)
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
        lak.percentage2.set("")
        lak.relation2.set("")
        lak.Name1.set("")
        lak.Name2.set("")
        lak.Name3.set("")
        lak.Name4.set("")
        lak.Name5.set("")
        lak.Name6.set("")
        lak.detail.set("")
        lak.percentage2.set(0)


if __name__== "__main__":
    root=Tk()
    aap= Bene_tool(root)
    root.mainloop()