# -*- coding: utf-8 -*-
"""
Created on Sat Feb  4 09:56:46 2023

@author: Basile"""

import locale
import pandas as pd
from icalendar import Calendar, Event
from datetime import datetime
import re
import hashlib
import tkinter as tk
from tkinter import ttk
import tkinter.filedialog as fd
import json


class Xls :
    def __init__(self):
        self.url='https://docs.google.com/spreadsheets/d/e/2PACX-1vT2LrtTfT7J-u9ueLq5hKFIz7HpUQoG3qj5FWklile7pg0Jo9msm9qLKk2x80Nhew/pubhtml?gid=1833983882&single=true&urp=gmail_link'
        self.date_format='%d %B %Y'
        self.df=None
          
    def acquire(self):
        # read the XLS file into a dataframe
        self.df = pd.read_html(self.url,encoding='UTF-8')
        self.df=self.df[0]
        
        # Supprimer les lignes/colonnes inutiles
        self.df.drop(index=[0,6,7,18,19,30,31,42,43,54,55,12,24,36,48,60],inplace=True)
        self.df.drop(index=self.df.index[50:],inplace=True)

        self.df.drop(columns=self.df.columns[0:3],inplace=True)

        
class Agenda :
    def __init__(self,xls,cours_suivis_list):
        self.cours_suivis_list=cours_suivis_list
        self.date_format='%Y%m%d'
        self.room_pattern=r'\d[A-Z]\d\d'
        self.course_type_pattern=r'BE(?=\D)|CM(?=\D)|TP(?=\D)|CM\d|BE\d|BE|TD\d|TD|TP\d|TP'
        self.events=[] # liste des evenements ical
        self.ical_events=[] # evenements ical
        self.horaires=[None,["080000","100000"],["080000","100000"],["101500","121500"],["101500","121500"],["133000","153000"],["133000","153000"],["154500","174500"],["154500","174500"]]
        self.nb_weeks=3 # number of weeks to add to calendar
        self.cours_suivis_pattern=re.compile("211-TP5|213-TD1|214-TD1|213-TP2|214-TP2|211-TD3|23\d-TP2|23\d-TD1")
        self.groupe_pattern=re.compile("gr\d|Gr\d")
        self.option=2 # =0 pour GM/1=GC/2=GE
        self.cours_option_pattern=[re.compile("211|212|213|22\d|20\d|Evènement particulier"),re.compile("211|212|213|24\d|20\d|Evènement particulier"),re.compile("211|213|214|23\d|20\d|Evènement particulier")]
        self.cal=None # ical Calendar
        self.df=xls.df # dataframe
        self.old_date_format='%d %B %Y'
        self.course_pattern=r'\d\d\d'
        self.checkbox=False


    def compare_uids(self,event): # compare les uid d'un event avec une liste d'events
        for e in self.events:
            if e['uid'] == event['uid']:
                return True
        return False
    
    def generate_cal(self):
        self.cal = Calendar()
        self.cal.add('prodid', '-//My calendar product//mxm.dk//')
        self.cal.add('version', '2.0')            
        for event in self.events :
            self.cal.add_component(event)    
        
    def to_ics(self, destination):
       # Create the .ics file
       with open(destination, 'wb') as f:
           f.write(self.cal.to_ical())
        
    def build(self):
        locale.setlocale(locale.LC_ALL, 'fr_FR.utf8') # change le format de date pour une date francaise
        imax=49 # nombre de lignes à prendre en compte apres xls.acquire()
        if self.option=="GM":
            self.option=1
        if self.option=="GC":
            self.option=0
        if self.option=="GE":
            self.option=2

        if self.nb_weeks>int((self.df.shape[1]-8)/12): # Si l'utilisateur a entrée un nombre de semaines trop grand
            self.nb_weeks=int((self.df.shape[1]-8)/12)
        for n in range(0,self.nb_weeks):
            for j in range(n*12,(n+1)*12):   
                i=1
                while i<imax:
                    if self.df.notna().iat[i,j] and (i%9!=0) and len(self.df.iat[i,j])!=1: # Si cellule non vide et ne correspond pas à une date
                        event=Event()
                        date=self.df.iat[(i//9)*9,(j//12)*12+6]
                        
                        #Changement de format de la date
                        date = datetime.strptime(date, self.old_date_format)
                        date = date.strftime(self.date_format)
            
                        # Ajout de la date/horaire de l'évènement
                        event['dtstart']=date + "T"+self.horaires[i%9][0]
                        event['dtend']=date +"T"+ self.horaires[i%9][1]
                        
                        # Extraction du texte 
                        event_str=str(self.df.iat[i,j])
                        i+=1
                        event_str+=" "+str(self.df.iat[i,j])

                        
                        # Gestion du type (BE1,TP2,...) de cours
                        suivi=False
                        cours=re.search(self.course_pattern, event_str).group() if re.search(self.course_pattern, event_str) else "Evènement particulier"
                        if re.fullmatch(self.cours_option_pattern[self.option],cours): # Si le cours correspond à la bonne filière
                            type_cours=re.findall(self.course_type_pattern, event_str) if re.findall(self.course_type_pattern, event_str) else ""
                            summary=cours+"-"
                            for type_cours_i in type_cours:
                                if type_cours_i not in summary :    
                                    summary+=type_cours_i+"/"
                            summary=summary[:-1] # enlève le dernier"/"

                            if re.fullmatch(re.compile("20\d"),cours):
                                suivi=True

                            if not self.checkbox:
                                if cours=="Evènement particulier":
                                    suivi=True
                                
                            if re.fullmatch(re.compile("21\d|23\d"),cours):
                                if type_cours=="": # Si le type de cours n'a pas été détecté
                                    suivi=True
                                else :
                                    if re.fullmatch(re.compile("CM\d"),type_cours[0]) or type_cours=="":
                                        suivi=True
                                    groupe=re.search(self.groupe_pattern, event_str).group() if re.search(self.groupe_pattern, event_str) else ""
                                    if groupe!="":
                                        groupe=groupe[-1]
                                        summary+=" GR" + groupe
                                    # le cours est-il suivi par l'utilisateur ?"
                                    if groupe=="": # Si le groupe n'a pas été détecté
                                        suivi=True
                                    for type_cours_i in type_cours:
                                        if re.fullmatch(self.cours_suivis_pattern, cours+type_cours_i[0:2]+groupe):
                                            suivi=True

                            if re.fullmatch(re.compile("24\d"),cours):
                                if type_cours=="": # Si le type de cours n'a pas été détecté
                                    suivi=True
                                else :
                                    type_cours=type_cours[0]
                                    if re.match(re.compile("CM"),type_cours):
                                        suivi=True
                                    groupe=type_cours[-1] if len(type_cours)==3 else "" # Si le type de cours est bien de la forme "BE2" ou "TD1" et non "CM" de
                                    if groupe!="":
                                        summary+=" GR" + groupe
                                    # le cours est-il suivi par l'utilisateur ?"
                                    if groupe=="": # Si le groupe n'a pas été détecté
                                        suivi=True
                                    if re.fullmatch(self.cours_suivis_pattern, cours+type_cours[0:2]+groupe):
                                        suivi=True


                            if re.match(re.compile("22\d"),cours): # Cours spécifiques aux GC
                                suivi=True

                        
                        # Ajout du titre/localisation/description
                        if suivi :    
                            salle=re.search(self.room_pattern,event_str).group() if re.search(self.room_pattern,event_str) else ""
                            event.add('summary',summary)
                            event.add('location',salle)
                            event.add('description',event_str)
                            
                            # Gestion du cas : cours de 4h
                            if self.df.iat[i+1,j]==self.df.iat[i,j]: # Si c'est un cours de 4h 1
                                event["dtend"]=date +"T"+ self.horaires[(i+1)%9][1]
                                i+=2 # On passe les deux créneaux qui suivent car correspondent au cours de 4h
                            
                            # Création d'une unique ID pour l'evenement
                            event_details = f"{event.get('summary')}-{event.get('dtstart')}-{event.get('dtend')}"
                            event['uid'] = hashlib.sha256(event_details.encode('utf-8')).hexdigest() 
                            
                            # Ajout de "event" à la liste d'évènements
                            
                            if not self.compare_uids(event): # Si l'évènement n'existe pas déjà
                                self.events.append(event)

                    # Gestion du cas : cours de 4h
                        if not suivi:
                            if self.df.iat[i+1,j]==self.df.iat[i,j]: # Si c'est un cours de 4h 2
                                i+=2 # On passe les deux créneaux qui suivent car correspondent au cours de 4h
                            

                    i+=1
            print("avancement : "+str(100*(n+1)/self.nb_weeks)+"%")

class GUI:
    def __init__(self,agenda):
        self.agenda=agenda
        self.add_button=None
        self.generate_button=None
        self.generate_button_label=None
        self.download_button=None
        self.download_button_label=None
        self.cours_suivis_label=None
        self.filiere=None
        self.filiere_label=None
        self.filieres=None
        self.filieres_option_menu1=None
        self.frame=None
        self.groupes_imput_field=None
        self.inner_frame1=None
        self.inner_frame2=None
        self.inner_frame3=None
        self.inner_frame4=None
        self.listbox=None
        self.modify_button=None
        self.nb_semaines=None
        self.nb_semaines_input_button=None
        self.nb_semaines_input_field=None
        self.nb_semaines_label=None
        self.nb_semaines_result_label=None
        self.remove_button=None
        self.root=None
        self.destionation=None
        self.test_button=None
        self.data={}
        self.creator_label=None
        self.checkbox_var=None
    
    def filiere_selected(self,value):
        self.agenda.option=value
        self.data["filieres_option_menu1"]=self.agenda.option

    def generate_button_clicked(self):
        cours_suivis_list = self.listbox.get(0, tk.END)
        self.agenda.cours_suivis_pattern=""
        # Collect data
        self.data["cours_suivis"]=[]
        for cours_suivi in cours_suivis_list:
            self.data["cours_suivis"].append(cours_suivi)
            if "x" in cours_suivi or "X" in cours_suivi:
                cours_suivi=re.sub("[xX]", r"\\d",cours_suivi)
                self.agenda.cours_suivis_pattern+=cours_suivi[0:4]+cours_suivi[4:8]+"|"
            else:
                self.agenda.cours_suivis_pattern+=cours_suivi[0:3]+cours_suivi[4:7]+"|"
        self.agenda.cours_suivis_pattern=re.compile(self.agenda.cours_suivis_pattern)
        self.data["checkbox"]= self.checkbox_var.get()==1
        self.agenda.checkbox= self.checkbox_var.get()==1
        self.save_data() # Save what user entered
        
        self.agenda.build()
        self.agenda.generate_cal()
        self.generate_button_label.config(text="Agenda créé")

    def download_button_clicked(self):
        self.destination = fd.asksaveasfilename(defaultextension='.ics')
        self.agenda.to_ics(self.destination)
        self.download_button_label.config(text="Téléchargé")
        

    def get_input_nb_semaines(self,event=None):
        nb_semaines = self.nb_semaines_input_field.get()
        self.agenda.nb_weeks=int(nb_semaines)
        self.data["nb_weeks"]=self.agenda.nb_weeks
        self.nb_semaines_result_label.config(text="Nombre de semaines : " + nb_semaines)
        
    def add_to_list(self,event=None):
        item = self.groupes_input_field.get()
        if item!="":
            self.listbox.insert(tk.END, item)
            self.groupes_input_field.delete(0, tk.END)

    def remove_from_list(self,event=None):
        index = self.listbox.curselection()
        if index:
            self.listbox.delete(index)

    def modify_list(self):
        index = self.listbox.curselection()
        if index:
            item = self.listbox.get(index)
            self.groupes_input_field.delete(0, tk.END)
            self.groupes_input_field.insert(0, item)
            self.listbox.delete(index)

    def save_data(self):
        # Save the data to a file
        with open('data.json', 'w') as f:
            json.dump(self.data, f)

    def load_data(self):
        # Load the data from the file
        try:
            with open('data.json', 'r') as f:
                self.data = json.load(f)
        except FileNotFoundError:
            # If the file doesn't exist, create a new empty dictionary
            self.data = {}
            with open('data.json', 'w') as f:
               json.dump(self.data, f)
        
    def build(self):
        self.root = tk.Tk()
        self.root.geometry("370x450")
        self.root.title("Saphire Agenda")

        self.frame = ttk.Frame(self.root, padding="10 10 10 10")
        self.frame.grid(column=0, row=0, sticky=(tk.N, tk.W, tk.E, tk.S))

      
        # Option menu des filières
        self.inner_frame1=ttk.Frame(self.frame,padding="2 2 2 2")
        self.inner_frame1.grid(column=0, row=0)
        self.filiere_label=tk.Label(self.inner_frame1,text="Option (GM,GC,GE) :")
        self.filiere_label.grid(column=0,row=0)
        self.filiere = tk.StringVar(self.frame)
        self.filiere.set("GM")
        self.filieres = ["option","GM", "GC", "GE"]
        self.filieres_option_menu1 = ttk.OptionMenu(self.inner_frame1, self.filiere, *self.filieres, command=self.filiere_selected)
        self.filiere.set(self.data.get("filieres_option_menu1"))
        self.agenda.option=self.data.get("filieres_option_menu1")
        self.filieres_option_menu1.grid(column=1,row=0)


        # Choix des cours suivis

        self.cours_suivis_label=tk.Label(self.frame,text="Entrer son groupe de TP/TD/BE.     Exemple :'232 BE1' ou '221 TP2'")
        self.cours_suivis_label.grid(column=0,row=1)

        self.groupes_input_field = ttk.Entry(self.frame)
        self.groupes_input_field.grid(column=0,row=2)
  

        # Widget pour les 3 bouttons
        self.inner_frame2=ttk.Frame(self.frame,padding="2 2 2 2")
        self.inner_frame2.grid(column=0, row=3)
        
        self.add_button = ttk.Button(self.inner_frame2, text="Ajouter", command=self.add_to_list)
        self.add_button.grid(column=0,row=0)
        self.groupes_input_field.bind("<Return>",self.add_to_list) # <Entrer> ajoute l'élément à la liste
        
        self.remove_button = ttk.Button(self.inner_frame2, text="Retirer", command=self.remove_from_list)
        self.remove_button.grid(column=1,row=0)
        
        
        self.modify_button = ttk.Button(self.inner_frame2, text="Modifier", command=self.modify_list)
        self. modify_button.grid(column=2,row=0)

        self.listbox = tk.Listbox(self.frame)
        # Default values for the listbox
        for cours_suivi in self.data.get("cours_suivis",""):
            self.listbox.insert(tk.END, cours_suivi)
        self.listbox.grid(column=0,row=4)
        self.listbox.bind("<Delete>",self.remove_from_list)

        # Choix du nombre de semaines
        self.nb_semaines_label = ttk.Label(self.frame,text="Nombre de semaines à ajouter à l'agenda.     Exemple :'3'\n (Nb de semaines sur l'emploi du temps excel)")
        self.nb_semaines_label.grid(column=0,row=5)
        self.inner_frame3=ttk.Frame(self.frame,padding="2 2 2 2")
        self.inner_frame3.grid(column=0,row=6)
        self.nb_semaines = tk.StringVar()
        self.nb_semaines_input_field = ttk.Entry(self.inner_frame3)
        self.nb_semaines_input_field.insert(0,self.data.get("nb_weeks",""))
        self.agenda.nb_weeks=self.data.get("nb_weeks")
        self.nb_semaines_input_field.grid(column=0,row=0)
        self.nb_semaines_input_field.bind("<Return>",self.get_input_nb_semaines)
        self.nb_semaines_result_label =tk.Label(self.frame,text="")

        self.nb_semaines_input_button = ttk.Button(self.inner_frame3, text="Confirmer", command=self.get_input_nb_semaines)
        self.nb_semaines_input_button.grid(column=1,row=0)

        self.nb_semaines_result_label.grid(column=0,row=7)



        self.inner_frame4=ttk.Frame(self.frame,padding="2 2 2 2")
        self.inner_frame4.grid(column=0,row=8)
        # Download button / generate button
        self.generate_button = ttk.Button(self.inner_frame4, text="Générer l'agenda (.ics)", command=self.generate_button_clicked)
        self.generate_button.grid(column=0,row=0)
        self.generate_button_label = tk.Label(self.inner_frame4, text="",foreground="green")
        self.generate_button_label.grid(column=0,row=1)

        self.download_button = ttk.Button(self.inner_frame4, text="Télécharger", command=self.download_button_clicked)
        self.download_button.grid(column=1,row=0)
        self.download_button_label = ttk.Label(self.inner_frame4, text="", foreground="green")
        self.download_button_label.grid(column=1,row=1)

        # Checkbox
        self.checkbox_var = tk.IntVar()
        if "checkbox" in self.data :
            self.checkbox_var.set(self.data["checkbox"])
        checkbox = ttk.Checkbutton(self.inner_frame4, text="Sans évènement\n particulier", variable=self.checkbox_var)
        checkbox.grid(column=2,row=0)
        

        # Creator label
        self.creator_label=tk.Label(self.frame, text="Créé par Basile PLUS", foreground="gray")
        self.creator_label.grid(column=0,row=12)
    
    def run(self):
        self.root.mainloop()
        
        
if __name__ == "__main__":      
    xls=Xls()
    xls.acquire()
    
    ag=Agenda(xls,[])

    gui=GUI(ag)
    gui.load_data()
    gui.build()
    gui.run()



        

