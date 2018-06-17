# -*- coding: utf-8 -*-
"""
Created on Mon Apr 30 19:58:11 2018

@author: Nicola
"""

import math
import sys
import numpy

def Encode(dada):
    try:
        if type(dada) == unicode:
            dada = dada.encode('utf-8')
            dada = dada.strip()
        if dada == 'NN' or dada == '(NN)' or dada == 'x' or dada == '':
            dada = None
        if (type(dada) == float or type(dada) == numpy.float64) and math.isnan(dada):
            dada = None
        if type(dada) == numpy.int64:
            dada = int(dada.item())
    except:
        e = sys.exc_info()[0]
        print(e)
        
    return dada


def insertPersonaBuidatRev(Persona_Buidat_rev,dada):
    dada = Encode(dada)
    if not any(e['codiNom']== dada for e in Persona_Buidat_rev) and dada != None and len(dada)<=4:
        newdict = {'codiNom' : dada}
        Persona_Buidat_rev = Persona_Buidat_rev + (newdict,)
    return Persona_Buidat_rev
    
def insertParroquia(Parroquia,dada):
    dada = Encode(dada)
    if not any(e['nomParroquia'] == dada for e in Parroquia) and dada != None:
        newdict = {'nomParroquia':dada}
        Parroquia = Parroquia + (newdict,)
    return Parroquia
    
def insertArxiu(Arxiu,nomllibre,nomparroquia,pbuidat,previsat):
    nomllibre = Encode(nomllibre)
    nomparroquia = Encode(nomparroquia)
    pbuidat = Encode(pbuidat)
    previsat = Encode(previsat)
    if not any(e['NomLLibre'] == nomllibre for e in Arxiu) and nomllibre != None:
        newdict = {'NomLLibre' : nomllibre, 'NomParroquia' : nomparroquia,
                   'PBuidat' : pbuidat, 'PRevisat': previsat}
        Arxiu = Arxiu + (newdict,)
    return Arxiu

def insertRegistre(Registre, nomllibre, numregistre,pgllibre,pgpdf):
    nomllibre = Encode(nomllibre)
    pgllibre = Encode(pgllibre)
    numregistre = Encode(numregistre)
    pgpdf = Encode(pgpdf)
    if type(pgpdf) != int: pgpdf = None
    if not any(e['NomLLibre'] == nomllibre and e['NumRegistre'] == numregistre for e in Registre) and numregistre != None:
        newdict = {'NomLLibre':nomllibre, 'NumRegistre':numregistre,'PgLlibre':pgllibre,
                   'PgPdf':pgpdf}
        Registre = Registre + (newdict,)
    return Registre

def insertEvent(Event, nomllibre, numregistre, datains, dataeve, lloceve, tipus, obsGen):
    nomllibre = Encode(nomllibre);numregistre = Encode(numregistre)
    datains = Encode(datains)
    dataeve = Encode(dataeve)
    if datains != None and len(datains)>12: datains=None
    if dataeve != None and len(dataeve)>12: dataeve=None
    lloceve = Encode(lloceve)
    obsGen = Encode(obsGen)
    if obsGen != None and len(obsGen) > 500 : obsGen = obsGen[:500]
    if not any(e['NomLLibre'] == nomllibre and e['NumRegistre'] == numregistre for e in Event) and numregistre != None:
        newdict = {'NomLLibre':nomllibre, 'NumRegistre':numregistre, 'DataIns':datains,
                   'DataEve':dataeve, 'LlocEve':lloceve, 'ObsGen':obsGen, 'Tipus':tipus}
        Event = Event + (newdict,)
    return Event

def insertPersona(Persona,nom,cognom1,cognom2,sexe,estat_vital,ofici_carrec,alies,estat_civil,
                  data_naixement, lloc_naixement, residencia):
    nom = Encode(nom);cognom1 = Encode(cognom1);cognom2= Encode(cognom2);sexe = Encode(sexe)
    estat_vital = Encode(estat_vital); ofici_carrec = Encode(ofici_carrec)
    alies = Encode(alies);estat_civil = Encode(estat_civil)
    data_naixement = Encode(data_naixement); lloc_naixement = Encode(lloc_naixement)
    residencia = Encode(residencia)
    if type(data_naixement) == int: data_naixement = str(data_naixement) 
    if data_naixement != None and len(data_naixement)>12: data_naixement=None
    if sexe != None: sexe = sexe.upper()
    if sexe != 'M' and sexe!= 'F' and sexe != 'A': sexe = None
    if estat_vital != None and (estat_vital == 'D' or 'difunt' in estat_vital): estat_vital = 'difunt/a'
    else: estat_vital = 'viu/va' 
    if estat_civil != None and len(estat_civil) > 50: estat_civil = estat_civil[:50]
    
    if cognom1 != None:
        if '+' in cognom1: estat_vital = 'difunt/a'
        if len(cognom1) >50: cognom1 = cognom1[:50]
    if nom != None and len(nom)>50: nom = nom[:50]
    if cognom2 != None and len(cognom2) >50: cognom2 = cognom2[:50]
    
    if not (nom == None and cognom1 == None and cognom2 == None) and not any(e['Nom'] == nom
           and e['Cognom1'] == cognom1 and e['Cognom2'] == cognom2 and e['Est_vital'] == estat_vital
           and e['Data_Naix'] == data_naixement and e['Lloc_Naix']== lloc_naixement for e in Persona):
        
        newdict = {'Nom' : nom, 'Cognom1': cognom1, 'Cognom2':cognom2, 'Sexe':sexe,
                   'Est_vital' : estat_vital, 'Of_Carrec':ofici_carrec, 'Alies':alies,
                   'Est_civil':estat_civil, 'Data_Naix':data_naixement,
                   'Lloc_Naix':lloc_naixement, 'Residencia':residencia}
        Persona = Persona + (newdict,)
    return Persona
        
def insertParticipant(Participant,nomllibre,numregistre,tipus,tipusfamilia,nom,cognom1,cognom2,
                      estat_vital,ofici_carrec,lloc_naixement, residencia, observacions):
    nomllibre = Encode(nomllibre); numregistre = Encode(numregistre)
    nom = Encode(nom);cognom1 = Encode(cognom1)
    cognom2 = Encode(cognom2);estat_vital = Encode(estat_vital);ofici_carrec = Encode(ofici_carrec)
    lloc_naixement = Encode(lloc_naixement); residencia = Encode(residencia)
    observacions = Encode(observacions)
    
    if observacions != None and len(observacions) > 500: observacions = observacions[:500]
    if estat_vital != None and (estat_vital == 'D' or 'difunt' in estat_vital): estat_vital = 'difunt/a'
    else: estat_vital = 'viu/va' 
    if cognom1 != None:
        if '+' in cognom1: estat_vital = 'difunt/a'
        if len(cognom1) >50: cognom1 = cognom1[:50]
    if nom != None and len(nom)>50: nom = nom[:50]
    if cognom2 != None and len(cognom2) >50: cognom2 = cognom2[:50]
    
    
    if not (nom == None and cognom1 == None and cognom2 == None):
        newdict = {'NomLLibre':nomllibre, 'NumRegistre':numregistre, 'Tipus':tipus,
                   'TipusFamilia': tipusfamilia,'Nom' : nom, 'Cognom1': cognom1,
                   'Cognom2':cognom2,'Est_vital' : estat_vital, 'Of_Carrec':ofici_carrec,
                   'Lloc_Naix':lloc_naixement, 'Residencia':residencia,'Obs':observacions}
        Participant = Participant + (newdict,)
    return Participant
    
def insertBaptisme(Baptisme,nomllibre,numregistre,nom,nomscomp,cognom1,cognom2,sexe,data_naixement,
                   lloc_naixement, observacio):
    nomllibre = Encode(nomllibre);numregistre = Encode(numregistre)
    nom = Encode(nom);cognom1 = Encode(cognom1)
    cognom2 = Encode(cognom2);nomscomp = Encode(nomscomp);sexe = Encode(sexe)
    data_naixement = Encode(data_naixement); lloc_naixement = Encode(lloc_naixement)
    observacio = Encode(observacio)
    
    if data_naixement != None and len(data_naixement)>12: data_naixement=None
    if observacio != None and len(observacio) > 500: observacio = observacio[:500]
    if sexe != None: sexe = sexe.upper()
    if sexe != 'M' and sexe!= 'F' and sexe != 'A': sexe = None
    if cognom1 != None and len(cognom1) >50: cognom1 = cognom1[:50]
    if nom != None and len(nom)>50: nom = nom[:50]
    if cognom2 != None and len(cognom2) >50: cognom2 = cognom2[:50]    
    
    if not (nom == None and cognom1 == None and cognom2 == None) and not any(e['NomLLibre'] == nomllibre and e['NumRegistre'] == numregistre for e in Baptisme) and numregistre != None:
        newdict = {'NomLLibre':nomllibre, 'NumRegistre':numregistre,'Nom' : nom,
                   'Noms_Comp':nomscomp,'Cognom1': cognom1, 'Cognom2':cognom2,
                   'Sexe':sexe,'Data_Naix':data_naixement,'Lloc_Naix':lloc_naixement,
                   'Obs':observacio}
        Baptisme = Baptisme + (newdict,)
    return Baptisme
    
def insertMatrimoni(Matrimoni,nomllibre,numregistre,nommarit,cognom1marit,cognom2marit,
                    edatmarit,lloc_naixementmarit,estat_civilmarit,residenciamarit,ocupaciomarit,
                    nommuller,cognom1muller,cognom2muller,aliesmuller,lloc_naixementmuller,
                    residenciamuller,edatmuller,estat_civilmuller,esglesia,capitols_mat):
    nomllibre = Encode(nomllibre);numregistre = Encode(numregistre)
    nommarit = Encode(nommarit);cognom1marit = Encode(cognom1marit)
    cognom2marit = Encode(cognom2marit);edatmarit = Encode(edatmarit)
    lloc_naixementmarit = Encode(lloc_naixementmarit); estat_civilmarit = Encode(estat_civilmarit)
    residenciamarit = Encode(residenciamarit); ocupaciomarit = Encode(ocupaciomarit)
    nommuller = Encode(nommuller); cognom1muller = Encode(cognom1muller); cognom2muller = Encode(cognom2muller)
    aliesmuller = Encode(aliesmuller);lloc_naixementmuller = Encode(lloc_naixementmuller)
    residenciamuller = Encode(residenciamuller);edatmuller = Encode(edatmuller)
    estat_civilmuller = Encode(estat_civilmuller); esglesia = Encode(esglesia); capitols_mat = Encode(capitols_mat)
    
    if cognom1marit != None and len(cognom1marit) >50: cognom1marit = cognom1marit[:50]
    if nommarit != None and len(nommarit)>50: nommarit = nommarit[:50]
    if cognom2marit != None and len(cognom2marit) >50: cognom2marit = cognom2marit[:50] 
    if cognom1muller != None and len(cognom1muller) >50: cognom1muller = cognom1muller[:50]
    if nommuller != None and len(nommuller)>50: nommuller = nommuller[:50]
    if cognom2muller != None and len(cognom2muller) >50: cognom2muller = cognom2muller[:50]    
    if estat_civilmarit != None and len(estat_civilmarit) > 50: estat_civilmarit = estat_civilmarit[:50]
    if estat_civilmuller != None and len(estat_civilmuller) > 50: estat_civilmuller = estat_civilmuller[:50]
    
    if not any(e['NomLLibre'] == nomllibre and e['NumRegistre'] == numregistre for e in Matrimoni) and numregistre != None:
        newdict = {'NomLLibre':nomllibre, 'NumRegistre':numregistre,'Nommarit':nommarit,
                   'Cognom1Marit':cognom1marit,'Cognom2Marit':cognom2marit,'EdatMarit':edatmarit,
                   'Lloc_Naix_Marit':lloc_naixementmarit,'Estat_civil_Marit':estat_civilmarit,
                   'ResidenciaMarit':residenciamarit,'OcupacioMarit':ocupaciomarit,
                    'NomMuller':nommuller,'Cognom1Muller':cognom1muller,'Cognom2Muller':cognom2muller,
                    'AliesMuller':aliesmuller,'Lloc_NaixementMuller':lloc_naixementmuller,
                    'ResidenciaMuller':residenciamuller,'EdatMuller':edatmuller,
                    'Estat_civilMuller':estat_civilmuller,'Esglesia':esglesia,
                    'Capitols_Mat':capitols_mat}
        Matrimoni = Matrimoni + (newdict,)
    return Matrimoni

def insertObituari(Obituari,nomllibre,numregistre,nom,cognom1,cognom2,alies,sexe,data_naixement,
                   lloc_naixement,edat,estat_civil,residencia,ocupacio,dataenterrament,lloc_enterrament,cementiri,
                   noms_fills,obs):
    nomllibre = Encode(nomllibre);numregistre = Encode(numregistre)
    nom = Encode(nom);cognom1 = Encode(cognom1);cognom2 = Encode(cognom2);
    alies = Encode(alies);sexe = Encode(sexe);data_naixement = Encode(data_naixement)
    lloc_naixement = Encode(lloc_naixement);edat = Encode(edat);estat_civil = Encode(estat_civil)
    residencia = Encode(residencia);ocupacio = Encode(ocupacio);lloc_enterrament = Encode(lloc_enterrament)
    cementiri = Encode(cementiri); noms_fills = Encode(noms_fills);obs = Encode(obs)
    dataenterrament = Encode(dataenterrament)
    
    if cognom1 != None and len(cognom1) >50: cognom1 = cognom1[:50]
    if nom != None and len(nom)>50: nom = nom[:50]
    if cognom2 != None and len(cognom2) >50: cognom2 = cognom2[:50]
    if type(data_naixement) == int: data_naixement = str(data_naixement)
    if data_naixement != None and len(data_naixement)>12: data_naixement = None
    if dataenterrament != None and len(dataenterrament)>12: dataenterrament = None
    if sexe != None: sexe = sexe.upper()
    if sexe != 'M' and sexe!= 'F' and sexe != 'A': sexe = None    
    
    if not (nom == None and cognom1 == None and cognom2 == None) and not any(e['NomLLibre'] == nomllibre and e['NumRegistre'] == numregistre for e in Obituari) and numregistre != None:
        newdict = {'NomLLibre':nomllibre, 'NumRegistre':numregistre, 'Nom':nom,
                   'Cognom1': cognom1, 'Cognom2':cognom2,'Alies':alies,'Sexe':sexe,
                   'Data_Naix':data_naixement,'Lloc_Naix':lloc_naixement,'Edat':edat,
                   'Estat_civil':estat_civil,'Residencia':residencia,'Ocupacio':ocupacio,
                   'Lloc_Enterrament':lloc_enterrament,'Cementiri':cementiri,'Noms_fills':noms_fills,
                   'Observacio':obs,'Data_Ent':dataenterrament}
        Obituari = Obituari +(newdict,)
    return Obituari

