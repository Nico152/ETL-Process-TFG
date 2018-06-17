# -*- coding: utf-8 -*-
"""
Created on Sat Apr 28 10:22:23 2018

@author: Nicola
"""
#Importem els moduls necessaris
import psycopg2
import os
import pandas as pd
import time


os.chdir('C:\Users\Nicola\Documents\UPC\Sistemes d Informacio\TFG\Altres')

#Importo el modul Inserts on hi ha les funcions que utilitzo per insertar a les
#variables de l'entorn que simulen les taules
import Inserts
Inserts = reload(Inserts)
os.chdir('C:\Users\Nicola\Documents\UPC\Sistemes d Informacio\TFG\Altres\Buidatge Agullana Total\Baptisme')
# Change directory 

# List all files and directories in current directory

listadirectorios = os.listdir('.')


#Creació taules a python
Persona_buidat_rev = ()
Parroquia = ()
Arxiu = ()
Registre = ()
Event = ()
Persona = ()
Baptisme = ()
Matrimoni = ()
Obituari = ()
Participant = ()

#%% -------- BAPTISMES -----------
#%%
count = 1;
# inicio la variable start_time per calcular el temps que trigui el proces
start_time = time.clock()
for file in listadirectorios:
    
    # Carrego un dels fitxers Excel
    xl = pd.ExcelFile(file)

    print(count,time.clock()-start_time)
    count +=1
    # Carrego la fulla 'baptismes' com a DataFrame
    df1 = xl.parse('baptismes')
    #Obtinc el nom de les columnes per accedir-hi després.
    cols = df1.columns
    #Recorro totes les fileres i creo una variable per a cada camp
    for i in range(0,len(df1.get(cols[0]))):
        NumRegistre = df1.get_value(i,cols[0])
        Persona_Buidatge = df1.get_value(i,cols[1])
        Persona_Rev = df1.get_value(i,cols[2])
        Nom_Parroquia = df1.get_value(i,cols[3])
        Nom_Llibre = df1.get_value(i,cols[4])
        Pag_Llibre = df1.get_value(i,cols[5])
        Pag_PDF = df1.get_value(i,cols[6])
        Data_Inscripcio = df1.get_value(i,cols[7])
        Nom_Oficiant = df1.get_value(i,cols[8])
        Cognoms_Oficiant = df1.get_value(i,cols[9])
        Ocupacio_Oficiant = df1.get_value(i,cols[10])
        Observacio_Oficiant = df1.get_value(i,cols[11])
        Nom_Bapt = df1.get_value(i,cols[12])
        Noms_Complementaris = df1.get_value(i,cols[13])
        Cognom1_Bapt = df1.get_value(i,cols[14])
        Cognom2_Bapt = df1.get_value(i,cols[15])
        Sexe_Bapt = df1.get_value(i,cols[16])
        Data_Naix = df1.get_value(i,cols[17])
        Lloc_Naix = df1.get_value(i,cols[18])
        Data_Bapt = df1.get_value(i,cols[19])
        Lloc_Bapt = df1.get_value(i,cols[20])
        Obs_Bapt = df1.get_value(i,cols[21])
        Nom_Pare = df1.get_value(i,cols[22])
        Cognom1_Pare = df1.get_value(i,cols[23])
        Cognom2_Pare = df1.get_value(i,cols[24])
        Lloc_Naix_Pare = df1.get_value(i,cols[25])
        Ocupacio_Pare = df1.get_value(i,cols[26])
        Lloc_Habit_Pare = df1.get_value(i,cols[27])
        Obs_Pare = df1.get_value(i,cols[28])
        Nom_Mare = df1.get_value(i,cols[29])
        Cognom1_Mare = df1.get_value(i,cols[30])
        Cognom2_Mare = df1.get_value(i,cols[31])
        Lloc_Naix_Mare = df1.get_value(i,cols[32])
        Obs_Mare = df1.get_value(i,cols[33])
        Nom_Padri = df1.get_value(i,cols[34])
        Cognom1_Padri = df1.get_value(i,cols[35])
        Cognom2_Padri = df1.get_value(i,cols[36])
        Lloc_Naix_Padri = df1.get_value(i,cols[37])
        Estat_Padri = df1.get_value(i,cols[38])
        Ocupacio_Padri = df1.get_value(i,cols[39])
        Lloc_Habit_Padri = df1.get_value(i,cols[40])
        Parentesc_Padri = df1.get_value(i,cols[41])
        Obs_Padri = df1.get_value(i,cols[42])
        Nom_Padrina = df1.get_value(i,cols[43])
        Cognom1_Padrina = df1.get_value(i,cols[44])
        Cognom2_Padrina = df1.get_value(i,cols[45])
        Lloc_Naix_Padrina = df1.get_value(i,cols[46])
        Estat_Padrina = df1.get_value(i,cols[47])
        Lloc_Habit_Padrina = df1.get_value(i,cols[48])
        Parentesc_Padrina = df1.get_value(i,cols[49])
        Obs_Padrina = df1.get_value(i,cols[53])
        Nom_Conj_Padrina = df1.get_value(i,cols[54])
        Cognom_Conj_Padrina = df1.get_value(i,cols[55])
        Ocupacio_Conj_Padrina = df1.get_value(i,cols[56])
        Observacio_Conj_Padrina = df1.get_value(i,cols[57])
        Nom_Avi_P = df1.get_value(i,cols[58])
        Cognom1_Avi_P = df1.get_value(i,cols[59])
        Cognom2_Avi_P = df1.get_value(i,cols[60])
        Lloc_Naix_Avi_P = df1.get_value(i,cols[61])
        Ocupacio_Avi_P = df1.get_value(i,cols[62])
        Lloc_Hab_Avi_P = df1.get_value(i,cols[63])
        Obs_Avi_P = df1.get_value(i,cols[64])
        Nom_Avia_P = df1.get_value(i,cols[65])
        Cognom1_Avia_P = df1.get_value(i,cols[66])
        Cognom2_Avia_P = df1.get_value(i,cols[67])
        Lloc_Naix_Avia_P = df1.get_value(i,cols[68]) 
        Lloc_Hab_Avia_P = df1.get_value(i,cols[69])
        Obs_Avia_P = df1.get_value(i,cols[70])
        Nom_Avi_M =  df1.get_value(i,cols[71])
        Cognom1_Avi_M = df1.get_value(i,cols[72])
        Cognom2_Avi_M = df1.get_value(i,cols[73])
        Lloc_Naix_Avi_M = df1.get_value(i,cols[74])
        Ocupacio_Avi_M = df1.get_value(i,cols[75])
        Lloc_Hab_Avi_M = df1.get_value(i,cols[76])
        Obs_Avi_M = df1.get_value(i,cols[77])   
        Nom_Avia_M = df1.get_value(i,cols[78])
        Cognom1_Avia_M = df1.get_value(i,cols[79])
        Cognom2_Avia_M = df1.get_value(i,cols[80])
        Lloc_Naix_Avia_M = df1.get_value(i,cols[81])
        Lloc_Hab_Avia_M = df1.get_value(i,cols[82])
        Obs_Avia_M = df1.get_value(i,cols[83])
        Observacions_Generals = df1.get_value(i,cols[84])
        
            
        #Inserto a cada taula els valors necessaris
        Persona_buidat_rev = Inserts.insertPersonaBuidatRev(Persona_buidat_rev,Persona_Buidatge)
        Persona_buidat_rev = Inserts.insertPersonaBuidatRev(Persona_buidat_rev,Persona_Rev)
        Parroquia = Inserts.insertParroquia(Parroquia,Nom_Parroquia)
        Arxiu = Inserts.insertArxiu(Arxiu,Nom_Llibre,Nom_Parroquia,Persona_Buidatge,Persona_Rev)
        Registre = Inserts.insertRegistre(Registre,Nom_Llibre,NumRegistre,Pag_Llibre,Pag_PDF)
        Event = Inserts.insertEvent(Event, Nom_Llibre,NumRegistre,Data_Inscripcio,Data_Bapt,Lloc_Bapt,
                                    'Baptisme',Observacions_Generals)
        
        #Oficiant
        Persona = Inserts.insertPersona(Persona,Nom_Oficiant,Cognoms_Oficiant,None,None,
                                        None,Ocupacio_Oficiant,None,None,None,None,None)
        Participant = Inserts.insertParticipant(Participant,Nom_Llibre,NumRegistre,'Baptisme',
                                                'Oficiant',Nom_Oficiant,Cognoms_Oficiant,None,
                                                None,Ocupacio_Oficiant,None,None,Observacio_Oficiant)
        
        #Protagonista
        Persona = Inserts.insertPersona(Persona,Nom_Bapt,Cognom1_Bapt,Cognom2_Bapt,Sexe_Bapt,None,
                                 None,None,None,Data_Naix,Lloc_Naix,None)
        Baptisme = Inserts.insertBaptisme(Baptisme,Nom_Llibre,NumRegistre,Nom_Bapt,Noms_Complementaris,
                                          Cognom1_Bapt,Cognom2_Bapt,Sexe_Bapt,Data_Naix,Lloc_Naix,Obs_Bapt)
        #Pare
        Persona = Inserts.insertPersona(Persona,Nom_Pare,Cognom1_Pare,Cognom2_Pare,'M',
                                        None,Ocupacio_Pare,None,None,None,Lloc_Naix_Pare,
                                        Lloc_Habit_Pare)
        Participant = Inserts.insertParticipant(Participant,Nom_Llibre,NumRegistre,'Baptisme',
                                                'Pare',Nom_Pare,Cognom1_Pare,Cognom2_Pare,
                                                None,Ocupacio_Pare,Lloc_Naix_Pare,
                                                Lloc_Habit_Pare,Obs_Pare)
        #Mare
        Persona = Inserts.insertPersona(Persona,Nom_Mare,Cognom1_Mare,Cognom2_Mare,'F',
                                        None,None,None,None,None,Lloc_Naix_Mare,
                                        None)
        Participant = Inserts.insertParticipant(Participant,Nom_Llibre,NumRegistre,'Baptisme',
                                                'Mare',Nom_Mare,Cognom1_Mare,Cognom2_Mare,
                                                None,None,Lloc_Naix_Mare,
                                                None,Obs_Mare)
        #Padri
        Persona = Inserts.insertPersona(Persona,Nom_Padri,Cognom1_Padri,Cognom2_Padri,'M',
                                        None,Ocupacio_Padri,None,Estat_Padri,None,Lloc_Naix_Padri,
                                        Lloc_Habit_Padri)
        Participant = Inserts.insertParticipant(Participant,Nom_Llibre,NumRegistre,'Baptisme',
                                                'Padri',Nom_Padri,Cognom1_Padri,Cognom2_Padri,
                                                None,Ocupacio_Padri,Lloc_Naix_Padri,
                                                Lloc_Habit_Padri,Obs_Padri)    
        #Padrina
        Persona = Inserts.insertPersona(Persona,Nom_Padrina,Cognom1_Padrina,Cognom2_Padrina,'F',
                                        None,None,None,Estat_Padrina,None,Lloc_Naix_Padrina,
                                        Lloc_Habit_Padrina)
        Participant = Inserts.insertParticipant(Participant,Nom_Llibre,NumRegistre,'Baptisme',
                                                'Padrina',Nom_Padrina,Cognom1_Padrina,Cognom2_Padrina,
                                                None,None,Lloc_Naix_Padrina,
                                                Lloc_Habit_Padrina,Obs_Padrina)    
        #Conjugue Padrina
        Persona = Inserts.insertPersona(Persona,Nom_Conj_Padrina,Cognom_Conj_Padrina,None,None,
                                        None,Ocupacio_Conj_Padrina,None,None,None,None,None)
        Participant = Inserts.insertParticipant(Participant,Nom_Llibre,NumRegistre,'Baptisme',
                                                'Conjugue Padrina',Nom_Conj_Padrina,Cognom_Conj_Padrina,None,
                                                None,Ocupacio_Conj_Padrina,None,None,Observacio_Conj_Padrina)
        #Avi Patern
        Persona = Inserts.insertPersona(Persona,Nom_Avi_P,Cognom1_Avi_P,Cognom2_Avi_P,'M',
                                        None,Ocupacio_Avi_P,None,None,None,Lloc_Naix_Avi_P,
                                        Lloc_Hab_Avi_P)
        Participant = Inserts.insertParticipant(Participant,Nom_Llibre,NumRegistre,'Baptisme',
                                                'Avi Patern',Nom_Avi_P,Cognom1_Avi_P,Cognom2_Avi_P,
                                                None,Ocupacio_Avi_P,Lloc_Naix_Avi_P,
                                                Lloc_Hab_Avi_P,Obs_Avi_P)
        #Avi Matern
        Persona = Inserts.insertPersona(Persona,Nom_Avi_M,Cognom1_Avi_M,Cognom2_Avi_M,'M',
                                        None,Ocupacio_Avi_M,None,None,None,Lloc_Naix_Avi_M,
                                        Lloc_Hab_Avi_M)
        Participant = Inserts.insertParticipant(Participant,Nom_Llibre,NumRegistre,'Baptisme',
                                                'Avi Matern',Nom_Avi_M,Cognom1_Avi_M,Cognom2_Avi_M,
                                                None,Ocupacio_Avi_M,Lloc_Naix_Avi_M,
                                                Lloc_Hab_Avi_M,Obs_Avi_M)
        #Avia Paterna
        Persona = Inserts.insertPersona(Persona,Nom_Avia_P,Cognom1_Avia_P,Cognom2_Avia_P,'M',
                                        None,None,None,None,None,Lloc_Naix_Avia_P,
                                        Lloc_Hab_Avia_P)
        Participant = Inserts.insertParticipant(Participant,Nom_Llibre,NumRegistre,'Baptisme',
                                                'Avia Paterna',Nom_Avia_P,Cognom1_Avia_P,Cognom2_Avia_P,
                                                None,None,Lloc_Naix_Avia_P,
                                                Lloc_Hab_Avia_P,Obs_Avia_P)
        #Avia Paterna
        Persona = Inserts.insertPersona(Persona,Nom_Avia_M,Cognom1_Avia_M,Cognom2_Avia_M,'M',
                                        None,None,None,None,None,Lloc_Naix_Avia_M,
                                        Lloc_Hab_Avia_M)
        Participant = Inserts.insertParticipant(Participant,Nom_Llibre,NumRegistre,'Baptisme',
                                                'Avia Materna',Nom_Avia_M,Cognom1_Avia_M,Cognom2_Avia_M,
                                                None,None,Lloc_Naix_Avia_M,
                                                Lloc_Hab_Avia_M,Obs_Avia_M)




#%%     --------------- MATRIMONIS ---------------------
#%%
os.chdir('C:\Users\Nicola\Documents\UPC\Sistemes d Informacio\TFG\Altres\Buidatge Agullana Total\Matrimoni')
# Change directory 

# List all files and directories in current directory
os.listdir('.')


listadirectorios = os.listdir('.')
count = 1;
start_time = time.clock()
for file in listadirectorios:
    # Load spreadsheet
    xl = pd.ExcelFile(file)
    
    # Print the sheet names
    print(count,time.clock()-start_time)
    count +=1
    # Load a sheet into a DataFrame by name: df1
    df1 = xl.parse('matrimonis')
    
    cols = df1.columns
    for i in range(0,len(df1.get(cols[0]))):
        NumRegistre = df1.get_value(i,cols[0])
        Persona_Buidatge = df1.get_value(i,cols[1])
        Persona_Rev = df1.get_value(i,cols[2])
        Nom_Parroquia = df1.get_value(i,cols[3])
        Nom_Llibre = df1.get_value(i,cols[4])
        Pag_Llibre = df1.get_value(i,cols[5])
        Pag_PDF = df1.get_value(i,cols[6])
        Data_Inscripcio = df1.get_value(i,cols[7])
        Nom_Oficiant = df1.get_value(i,cols[8])
        Cognoms_Oficiant = df1.get_value(i,cols[9])
        Ocupacio_Oficiant = df1.get_value(i,cols[10])
        Observacio_Oficiant = df1.get_value(i,cols[11])
        Lloc_Matrimoni = df1.get_value(i,cols[12])
        Esglesia = df1.get_value(i,cols[13])
        Data_Matr = df1.get_value(i,cols[14])
        Capt_Matrimonials = df1.get_value(i,cols[15])
        Nom_Marit = df1.get_value(i,cols[16])
        Cognom1_Marit = df1.get_value(i,cols[17])
        Cognom2_Marit = df1.get_value(i,cols[18])
        Lloc_Naix_Marit = df1.get_value(i,cols[19])
        Lloc_Habit_Marit = df1.get_value(i,cols[20])
        Edat_Marit = df1.get_value(i,cols[21])
        Est_civil_Marit = df1.get_value(i,cols[22])
        Ocupacio_Marit = df1.get_value(i,cols[23])
        Obs_Marit = df1.get_value(i,cols[24])
        Nom_Pare_Marit = df1.get_value(i,cols[25])
        Cognom1_Pare_Marit = df1.get_value(i,cols[26])
        Cognom2_Pare_Marit = df1.get_value(i,cols[27])
        Est_vital_Pare_Marit = df1.get_value(i,cols[28])
        Lloc_Naix_Pare_Marit = df1.get_value(i,cols[29])
        Lloc_Habit_Pare_Marit = df1.get_value(i,cols[30])
        Ocupacio_Pare_Marit = df1.get_value(i,cols[31])
        Obs_Pare_Marit = df1.get_value(i,cols[32])
        Nom_Mare_Marit = df1.get_value(i,cols[33])
        Cognom1_Mare_Marit = df1.get_value(i,cols[34])
        Cognom2_Mare_Marit = df1.get_value(i,cols[35])
        Est_vital_Mare_Marit = df1.get_value(i,cols[36])
        Lloc_Naix_Mare_Marit = df1.get_value(i,cols[37])
        Lloc_Habit_Mare_Marit = df1.get_value(i,cols[38])
        Obs_Mare_Marit = df1.get_value(i,cols[39])
        Nom_Muller = df1.get_value(i,cols[40])
        Cognom1_Muller = df1.get_value(i,cols[41])
        Cognom2_Muller = df1.get_value(i,cols[42])
        Alies_Muller = df1.get_value(i,cols[43])
        Data_Naix_Muller = df1.get_value(i,cols[44])
        Lloc_Naix_Muller = df1.get_value(i,cols[45])
        Edat_Muller = df1.get_value(i,cols[46])
        Est_civil_Muller = df1.get_value(i,cols[47])
        Lloc_Habit_Muller = df1.get_value(i,cols[48])
        Obs_Muller = df1.get_value(i,cols[49])
        Nom_Pare_Muller = df1.get_value(i,cols[50])
        Cognom1_Pare_Muller = df1.get_value(i,cols[51])
        Cognom2_Pare_Muller = df1.get_value(i,cols[52])
        Est_vital_Pare_Muller = df1.get_value(i,cols[53])
        Lloc_Naix_Pare_Muller = df1.get_value(i,cols[54])
        Lloc_Habit_Pare_Muller = df1.get_value(i,cols[55])
        Ocupacio_Pare_Muller = df1.get_value(i,cols[56])
        Obs_Pare_Muller = df1.get_value(i,cols[57])
        Nom_Mare_Muller = df1.get_value(i,cols[58])
        Cognom1_Mare_Muller = df1.get_value(i,cols[59])
        Cognom2_Mare_Muller = df1.get_value(i,cols[60])
        Est_vital_Mare_Muller = df1.get_value(i,cols[61])
        Lloc_Naix_Mare_Muller = df1.get_value(i,cols[62])
        Lloc_Habit_Mare_Muller = df1.get_value(i,cols[63])
        Obs_Mare_Muller = df1.get_value(i,cols[64])
        Nom_T_1 = df1.get_value(i,cols[65]) 
        Cognom_T_1 = df1.get_value(i,cols[66])
        Sexe_T_1 = df1.get_value(i,cols[67])
        Obs_T_1 =  df1.get_value(i,cols[68])
        Nom_T_2 = df1.get_value(i,cols[69])
        Cognom_T_2 = df1.get_value(i,cols[70])
        Sexe_T_2 = df1.get_value(i,cols[71])
        Obs_T_2 = df1.get_value(i,cols[72])
        Nom_T_3 = df1.get_value(i,cols[73])
        Cognom_T_3 = df1.get_value(i,cols[74])   
        Sexe_T_3 = df1.get_value(i,cols[75])
        Obs_T_3 = df1.get_value(i,cols[76])
        Observacions_Generals = df1.get_value(i,cols[77])
        
            
        
        Persona_buidat_rev = Inserts.insertPersonaBuidatRev(Persona_buidat_rev,Persona_Buidatge)
        Persona_buidat_rev = Inserts.insertPersonaBuidatRev(Persona_buidat_rev,Persona_Rev)
        Parroquia = Inserts.insertParroquia(Parroquia,Nom_Parroquia)
        Arxiu = Inserts.insertArxiu(Arxiu,Nom_Llibre,Nom_Parroquia,Persona_Buidatge,Persona_Rev)
        Registre = Inserts.insertRegistre(Registre,Nom_Llibre,NumRegistre,Pag_Llibre,Pag_PDF)
        Event = Inserts.insertEvent(Event, Nom_Llibre,NumRegistre,Data_Inscripcio,Data_Matr,
                                    Lloc_Matrimoni,'Matrimoni',Observacions_Generals)
        
        #Oficiant
        Persona = Inserts.insertPersona(Persona,Nom_Oficiant,Cognoms_Oficiant,None,None,
                                        None,Ocupacio_Oficiant,None,None,None,None,None)
        Participant = Inserts.insertParticipant(Participant,Nom_Llibre,NumRegistre,'Matrimoni',
                                                'Oficiant',Nom_Oficiant,Cognoms_Oficiant,None,
                                                None,Ocupacio_Oficiant,None,None,Observacio_Oficiant)
        
        #Matrimoni
        Matrimoni = Inserts.insertMatrimoni(Matrimoni,Nom_Llibre,NumRegistre,Nom_Marit,Cognom1_Marit,
                                            Cognom2_Marit,Edat_Marit,Lloc_Naix_Marit,Est_civil_Marit,
                                            Lloc_Habit_Marit,Ocupacio_Marit,Nom_Muller,Cognom1_Muller,
                                            Cognom2_Muller,Alies_Muller,Lloc_Naix_Muller,Lloc_Habit_Muller,
                                            Edat_Muller,Est_civil_Muller,Esglesia,Capt_Matrimonials)
        Persona = Inserts.insertPersona(Persona,Nom_Marit,Cognom1_Marit,Cognom2_Marit,'M',None,Ocupacio_Marit,
                                        None,Est_civil_Marit,None,Lloc_Naix_Marit,Lloc_Habit_Marit)
        Persona = Inserts.insertPersona(Persona,Nom_Muller,Cognom1_Muller,Cognom2_Muller,'F',None,None,
                                        Alies_Muller,Est_civil_Muller,Data_Naix_Muller,Lloc_Naix_Muller,Lloc_Habit_Muller)
        
        # Pare Marit
        Persona = Inserts.insertPersona(Persona,Nom_Pare_Marit,Cognom1_Pare_Marit,Cognom2_Pare_Marit,'M',Est_vital_Pare_Marit,Ocupacio_Pare_Marit,
                                        None,None,None,Lloc_Naix_Pare_Marit,Lloc_Habit_Pare_Marit)
        Participant = Inserts.insertParticipant(Participant,Nom_Llibre,NumRegistre,'Matrimoni',
                                                'Pare Marit',Nom_Pare_Marit,Cognom1_Pare_Marit,Cognom2_Pare_Marit,
                                                Est_vital_Pare_Marit,Ocupacio_Pare_Marit,Lloc_Naix_Pare_Marit,
                                                Lloc_Habit_Pare_Marit,Obs_Pare_Marit)
        
        # Pare Muller
        Persona = Inserts.insertPersona(Persona,Nom_Pare_Muller,Cognom1_Pare_Muller,Cognom2_Pare_Muller,'M',Est_vital_Pare_Muller,Ocupacio_Pare_Muller,
                                        None,None,None,Lloc_Naix_Pare_Muller,Lloc_Habit_Pare_Muller)
        Participant = Inserts.insertParticipant(Participant,Nom_Llibre,NumRegistre,'Matrimoni',
                                                'Pare Muller',Nom_Pare_Muller,Cognom1_Pare_Muller,Cognom2_Pare_Muller,
                                                Est_vital_Pare_Muller,Ocupacio_Pare_Muller,Lloc_Naix_Pare_Muller,
                                                Lloc_Habit_Pare_Muller,Obs_Pare_Muller)
        # Mare Marit
        Persona = Inserts.insertPersona(Persona,Nom_Mare_Marit,Cognom1_Mare_Marit,Cognom2_Mare_Marit,'M',Est_vital_Mare_Marit,None,
                                        None,None,None,Lloc_Naix_Mare_Marit,Lloc_Habit_Mare_Marit)
        Participant = Inserts.insertParticipant(Participant,Nom_Llibre,NumRegistre,'Matrimoni',
                                                'Mare Marit',Nom_Mare_Marit,Cognom1_Mare_Marit,Cognom2_Mare_Marit,
                                                Est_vital_Mare_Marit,None,Lloc_Naix_Mare_Marit,
                                                Lloc_Habit_Mare_Marit,Obs_Mare_Marit)
        # Mare Muller
        Persona = Inserts.insertPersona(Persona,Nom_Mare_Muller,Cognom1_Mare_Muller,Cognom2_Mare_Muller,'M',Est_vital_Mare_Muller,None,
                                        None,None,None,Lloc_Naix_Mare_Muller,Lloc_Habit_Mare_Muller)
        Participant = Inserts.insertParticipant(Participant,Nom_Llibre,NumRegistre,'Matrimoni',
                                                'Mare Muller',Nom_Mare_Muller,Cognom1_Mare_Muller,Cognom2_Mare_Muller,
                                                Est_vital_Mare_Muller,None,Lloc_Naix_Pare_Muller,
                                                Lloc_Habit_Mare_Muller,Obs_Mare_Muller)
        #Testimonis
        Persona = Inserts.insertPersona(Persona,Nom_T_1,Cognom_T_1,None,Sexe_T_1,None,None,None,None,None,None,None)
        Participant = Inserts.insertParticipant(Participant,Nom_Llibre,NumRegistre,'Matrimoni','Testimoni 1',
                                                Nom_T_1,Cognom_T_1,None,None,None,None,None,None) 
        Persona = Inserts.insertPersona(Persona,Nom_T_2,Cognom_T_2,None,Sexe_T_2,None,None,None,None,None,None,None)
        Participant = Inserts.insertParticipant(Participant,Nom_Llibre,NumRegistre,'Matrimoni','Testimoni 2',
                                                Nom_T_2,Cognom_T_2,None,None,None,None,None,None) 
        Persona = Inserts.insertPersona(Persona,Nom_T_3,Cognom_T_3,None,Sexe_T_3,None,None,None,None,None,None,None)
        Participant = Inserts.insertParticipant(Participant,Nom_Llibre,NumRegistre,'Matrimoni','Testimoni 3',
                                                Nom_T_3,Cognom_T_3,None,None,None,None,None,None) 

#######




# -------------- OBITUARIS ------------------------

os.chdir('C:\Users\Nicola\Documents\UPC\Sistemes d Informacio\TFG\Altres\Buidatge Agullana Total\Obituari')
# Change directory 

# List all files and directories in current directory
os.listdir('.')

listadirectorios = os.listdir('.')
count = 1;
start_time = time.clock()
for file in listadirectorios:
    # Load spreadsheet
    xl = pd.ExcelFile(file)
    
    # Print the sheet names
    print(count,time.clock()-start_time)
    count +=1
    # Load a sheet into a DataFrame by name: df1
    df1 = xl.parse('defuncions')
    
    cols = df1.columns
    for i in range(0,len(df1.get(cols[0]))):
        NumRegistre = df1.get_value(i,cols[0])
        Persona_Buidatge = df1.get_value(i,cols[1])
        Persona_Rev = df1.get_value(i,cols[2])
        Nom_Parroquia = df1.get_value(i,cols[3])
        Nom_Llibre = df1.get_value(i,cols[4])
        Pag_Llibre = df1.get_value(i,cols[5])
        Pag_PDF = df1.get_value(i,cols[6])
        Data_Inscripcio = df1.get_value(i,cols[7])
        Nom_Oficiant = df1.get_value(i,cols[8])
        Cognoms_Oficiant = df1.get_value(i,cols[9])
        Ocupacio_Oficiant = df1.get_value(i,cols[10])
        Observacio_Oficiant = df1.get_value(i,cols[11])
        Nom_Obi = df1.get_value(i,cols[12])
        Cognom1_Obi = df1.get_value(i,cols[13])
        Cognom2_Obi = df1.get_value(i,cols[14])
        Alies_Obi = df1.get_value(i,cols[15])
        Sexe_Obi = df1.get_value(i,cols[16])
        Data_Naix_Obi = df1.get_value(i,cols[17])
        Lloc_Naix_Obi = df1.get_value(i,cols[18])
        Edat_Obi = df1.get_value(i,cols[19])
        Estat_civil_Obi = df1.get_value(i,cols[20])
        Lloc_Habit_Obi = df1.get_value(i,cols[21])
        Ocupacio_Obi = df1.get_value(i,cols[22])
        Data_Def = df1.get_value(i,cols[23])
        Lloc_Def = df1.get_value(i,cols[24])
        Causa_Def = df1.get_value(i,cols[25])
        Data_Enterrament = df1.get_value(i,cols[26])
        Lloc_Enterrament = df1.get_value(i,cols[27])
        Cementiri = df1.get_value(i,cols[28])
        Obs_Obi = df1.get_value(i,cols[29])
        Nom_Conjugue_Obi = df1.get_value(i,cols[30])
        Cognom_Conjugue_Obi = df1.get_value(i,cols[31])
        Lloc_Naix_Conjugue_Obi = df1.get_value(i,cols[32])
        Lloc_Habit_Conjugue_Obi = df1.get_value(i,cols[33])
        Obs_Conjugue_Obi = df1.get_value(i,cols[34])
        Noms_fills = df1.get_value(i,cols[35])
        Nom_Pare_Obi = df1.get_value(i,cols[36])
        Cognom_Pare_Obi = df1.get_value(i,cols[37])
        Lloc_Naix_Pare_Obi = df1.get_value(i,cols[38])
        Ocupacio_Pare_Obi = df1.get_value(i,cols[39])
        Lloc_Habit_Pare_Obi = df1.get_value(i,cols[40])
        Obs_Pare_Obi = df1.get_value(i,cols[41])
        Nom_Mare_Obi = df1.get_value(i,cols[42])
        Cognom_Mare_Obi = df1.get_value(i,cols[43])
        Lloc_Naix_Mare_Obi = df1.get_value(i,cols[44])
        Lloc_Habit_Mare_Obi = df1.get_value(i,cols[45])
        Obs_Mare_Obi = df1.get_value(i,cols[46])
        Observacions_Generals = df1.get_value(i,cols[47])
        
            
        
        Persona_buidat_rev = Inserts.insertPersonaBuidatRev(Persona_buidat_rev,Persona_Buidatge)
        Persona_buidat_rev = Inserts.insertPersonaBuidatRev(Persona_buidat_rev,Persona_Rev)
        Parroquia = Inserts.insertParroquia(Parroquia,Nom_Parroquia)
        Arxiu = Inserts.insertArxiu(Arxiu,Nom_Llibre,Nom_Parroquia,Persona_Buidatge,Persona_Rev)
        Registre = Inserts.insertRegistre(Registre,Nom_Llibre,NumRegistre,Pag_Llibre,Pag_PDF)
        Event = Inserts.insertEvent(Event, Nom_Llibre,NumRegistre,Data_Inscripcio,Data_Def,
                                    Lloc_Def,'Obituari',Observacions_Generals)
        
        #Oficiant
        Persona = Inserts.insertPersona(Persona,Nom_Oficiant,Cognoms_Oficiant,None,None,
                                        None,Ocupacio_Oficiant,None,None,None,None,None)
        Participant = Inserts.insertParticipant(Participant,Nom_Llibre,NumRegistre,'Obituari',
                                                'Oficiant',Nom_Oficiant,Cognoms_Oficiant,None,
                                                None,Ocupacio_Oficiant,None,None,Observacio_Oficiant)
        # Obituari
        Persona = Inserts.insertPersona(Persona,Nom_Obi,Cognom1_Obi,Cognom2_Obi,Sexe_Obi,'difunt',
                                        Ocupacio_Obi,Alies_Obi,Estat_civil_Obi,Data_Naix_Obi,Lloc_Naix_Obi,
                                        Lloc_Habit_Obi)
        Obituari = Inserts.insertObituari(Obituari,Nom_Llibre,NumRegistre,Nom_Obi,Cognom1_Obi,Cognom2_Obi,
                                          Alies_Obi,Sexe_Obi,Data_Naix_Obi,Lloc_Naix_Obi,Edat_Obi,
                                          Estat_civil_Obi,Lloc_Habit_Obi,Ocupacio_Obi,Data_Enterrament,
                                          Lloc_Enterrament,Cementiri,Noms_fills,Obs_Obi)
        #Conjugue Obi
        Persona = Inserts.insertPersona(Persona,Nom_Conjugue_Obi,Cognom_Conjugue_Obi,None,None,None,None,
                                        None,'vidu/a',None,Lloc_Naix_Conjugue_Obi,Lloc_Habit_Conjugue_Obi)
        Participant = Inserts.insertParticipant(Participant,Nom_Llibre,NumRegistre,'Obituari','Conjugue Difunt',
                                                Nom_Conjugue_Obi,Cognom_Conjugue_Obi,None,None,None,
                                                Lloc_Naix_Conjugue_Obi,Lloc_Habit_Conjugue_Obi,Obs_Conjugue_Obi)
        # Pare Difunt
        Persona = Inserts.insertPersona(Persona,Nom_Pare_Obi,Cognom_Pare_Obi,None,'M',None,Ocupacio_Pare_Obi,
                                        None,None,None,Lloc_Naix_Pare_Obi,Lloc_Habit_Pare_Obi)
        Participant = Inserts.insertParticipant(Participant,Nom_Llibre,NumRegistre,'Obituari','Pare Difunt',
                                                Nom_Pare_Obi,Cognom_Pare_Obi,None,None,Ocupacio_Pare_Obi,
                                                Lloc_Naix_Pare_Obi,Lloc_Habit_Pare_Obi,Obs_Pare_Obi)
        # Mare Difunt
        Persona = Inserts.insertPersona(Persona,Nom_Mare_Obi,Cognom_Mare_Obi,None,'F',None,None,
                                        None,None,None,Lloc_Naix_Mare_Obi,Lloc_Habit_Mare_Obi)
        Participant = Inserts.insertParticipant(Participant,Nom_Llibre,NumRegistre,'Obituari','Mare Difunt',
                                                Nom_Mare_Obi,Cognom_Mare_Obi,None,None,None,
                                                Lloc_Naix_Mare_Obi,Lloc_Habit_Mare_Obi,Obs_Mare_Obi)




print(time.clock()-start_time)
try:
    conn = psycopg2.connect("dbname='TFG' user='postgres' host='localhost' password='peca2012'")
except:
    print "I am unable to connect to the database"

cur = conn.cursor()

cur.executemany("""INSERT INTO persona_buidatrev(codi_nom) VALUES (%(codiNom)s)""", Persona_buidat_rev)
cur.executemany("""INSERT INTO parroquia(nom_Parroquia) VALUES (%(nomParroquia)s)""", Parroquia)
cur.executemany("""INSERT INTO arxiu(nom_llibre,nom_Parroquia,p_buidat,p_revisat) VALUES (%(NomLLibre)s,%(NomParroquia)s,%(PBuidat)s,%(PRevisat)s)""",Arxiu)
cur.executemany("""INSERT INTO registre(nom_llibre,num_registre,pag_llibre,pag_pdf) VALUES (%(NomLLibre)s,%(NumRegistre)s,%(PgLlibre)s,%(PgPdf)s)""",Registre)
cur.executemany("""INSERT INTO event_taula(nom_llibre,num_registre,data_inscripcio,data_event,lloc_event,tipus_event,observaciogeneral) VALUES (%(NomLLibre)s,%(NumRegistre)s,%(DataIns)s,%(DataEve)s,%(LlocEve)s,%(Tipus)s,%(ObsGen)s)""",Event)
cur.executemany("""INSERT INTO persona(nom,cognom1,cognom2,sexe,estat_vital,ofici_carrec,alies,estat_civil,data_naixement,lloc_naixement,residencia) VALUES(%(Nom)s,%(Cognom1)s,%(Cognom2)s,%(Sexe)s,%(Est_vital)s,%(Of_Carrec)s,%(Alies)s,%(Est_civil)s,%(Data_Naix)s,%(Lloc_Naix)s,%(Residencia)s)""",Persona)
cur.executemany("""INSERT INTO baptisme(nom_llibre,num_registre,nom,noms_complementaris,cognom1,cognom2,sexe,data_naixement,lloc_naixement,observacio) VALUES(%(NomLLibre)s,%(NumRegistre)s,%(Nom)s,%(Noms_Comp)s,%(Cognom1)s,%(Cognom2)s,%(Sexe)s,%(Data_Naix)s,%(Lloc_Naix)s,%(Obs)s)""",Baptisme)
cur.executemany("""INSERT INTO matrimoni(nom_llibre,num_registre,nom_Marit,Cognom1_Marit,Cognom2_Marit,edat_Marit,lloc_naixement_Marit,estat_civil_Marit,residencia_Marit,ocupacio_Marit,nom_Muller,cognom1_Muller,cognom2_Muller,alies_Muller,lloc_naixement_Muller,residencia_Muller,edat_Muller,estat_civil_Muller,esglesia,capitols_Mat) VALUES(%(NomLLibre)s,%(NumRegistre)s,%(Nommarit)s,%(Cognom1Marit)s,%(Cognom2Marit)s,%(EdatMarit)s,%(Lloc_Naix_Marit)s,%(Estat_civil_Marit)s,%(ResidenciaMarit)s,%(OcupacioMarit)s,%(NomMuller)s,%(Cognom1Muller)s,%(Cognom2Muller)s,%(AliesMuller)s,%(Lloc_NaixementMuller)s,%(ResidenciaMuller)s,%(EdatMuller)s,%(Estat_civilMuller)s,%(Esglesia)s,%(Capitols_Mat)s)""",Matrimoni)
cur.executemany("""INSERT INTO obituari(nom_llibre,num_registre,nom,cognom1,cognom2,alies,sexe,data_naixement,lloc_naixement,edat,estat_civil,residencia,ocupacio,data_enterrament,lloc_enterrament,cementiri,noms_fills,Obs_Obi) VALUES(%(NomLLibre)s,%(NumRegistre)s,%(Nom)s,%(Cognom1)s,%(Cognom2)s,%(Alies)s,%(Sexe)s,%(Data_Naix)s,%(Lloc_Naix)s,%(Edat)s,%(Estat_civil)s,%(Residencia)s,%(Ocupacio)s,%(Data_Ent)s,%(Lloc_Enterrament)s,%(Cementiri)s,%(Noms_fills)s,%(Observacio)s)""",Obituari)
cur.executemany("""INSERT INTO participant(nom_llibre,num_registre,tipus_event,tipus_part,nom,cognom1,cognom2,estat_vital,ofici_carrec,lloc_naixement,residencia,observacio) VALUES(%(NomLLibre)s,%(NumRegistre)s,%(Tipus)s,%(TipusFamilia)s,%(Nom)s,%(Cognom1)s,%(Cognom2)s,%(Est_vital)s,%(Of_Carrec)s,%(Lloc_Naix)s,%(Residencia)s,%(Obs)s)""",Participant)
conn.commit()

cur.close()
conn.close()



#for e in Obituari:
#    if e['Lloc_Naix'] != None and len(e['Lloc_Naix']) > 100: print(e['Lloc_Naix'],'Lloc_Naix',len(e['Lloc_Naix']))
#    if e['Noms_fills'] != None and len(e['Noms_fills']) > 100: print(e['Noms_fills'],'Noms_fills',len(e['Noms_fills']))
#    if e['Cementiri'] != None and len(e['Cementiri']) > 100: print(e['Cementiri'],'Cementiri',len(e['Cementiri']))
#    