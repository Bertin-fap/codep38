from pathlib import Path
from datetime import date
import pandas as pd
import os
import json
from statistics import mean
from textwrap import wrap


import matplotlib.pyplot as plt
import numpy as np
pd.options.mode.chained_assignment = None 

def date_today():
    today = date.today()
    d = str(today.year)+str(today.month).zfill(2)+str(today.day).zfill(2)
    return d

def nbr_point_vote(n):
    bins = [2,11,21,36,51,76,101,151,201,251,301,401,501,601,701,801,901,1001,1501,2001,2501]
    return np.digitize(n,bins)

def dic_efv(root):
    # Extrait du sit EFV structue EXtraction/Autres rapports/ Ecole françaises de Vélo
    
    df_efv = pd.read_excel(root / Path('fichiers FFCT') / Path('EFV.xlsx'),skiprows=1)
    efv_dict = dict(zip(df_efv['Numéro Club'],df_efv['Président']))
    return efv_dict

def normalize_num_tel(s):
        s= str(s)
        if len(s)>0:
            digits = ''.join(char for char in s if char.isdigit())
            digits = digits.zfill(10)
            return " ".join(wrap(digits, width=2))
        else:
            return ''

def make_dic_age(root, year):

    efv_dict = dic_efv(root)
    efv_list = efv_dict.keys()

    effectif_list_h = []
    effectif_list_f = []
    dic_effectif = {}
    for file in os.listdir(root_effectif):
        df = pd.read_excel(root_effectif / Path(file))
        df['Date de naissance'] = pd.to_datetime(df['Date de naissance'],
                                                     format="%d/%m/%Y")
        df['Age']  = df['Date de naissance'].apply(lambda x :
                                                (pd.Timestamp(int(year), 9, 30)-x).days/365)
        name_club = int(file.split('.')[0])
        efv = 'oui' if name_club in efv_list else 'non'
        dic_effectif[name_club]= [name_club,
                                       round(df['Age'].mean(),2),
                                       round(df.query('Sexe =="Masculin"')['Age'].mean(),2),
                                       round(df.query('Sexe =="Féminin"')['Age'].mean(),2),
                                       round(df.query('Age >= 18')['Age'].mean(),2),
                                       round(df.query('Age < 18')['Age'].mean(),2),      
                                       len(df.query('Sexe =="Masculin"')),
                                       len(df.query('Sexe =="Masculin" & Age >= 18')),
                                       len(df.query('Sexe =="Féminin"')),
                                       len(df.query('Sexe =="Féminin" & Age >= 18')),
                                       round(100*len(df.query('Sexe =="Féminin"'))/len(df),2),
                                       len(df),efv]
        effectif_list_h.extend(df.query('Sexe =="Masculin"')['Age'].tolist())
        effectif_list_f.extend(df.query('Sexe =="Féminin"')['Age'].tolist())
    return dic_effectif,effectif_list_h, effectif_list_f
    
def correction_email_president(df,root):
    with open(root / Path('fichiers correctifs') / Path('correction_courriel_president.json'), 'r') as file:
        dic_correction = json.load(file)
    for num_licence in dic_correction.keys():
            idx = df[df["N°"]==int(num_licence)].index
            df.loc[idx,'Adresse email'] = dic_correction[num_licence]
    return df
    
def correction_email_secretaire(df):
    with open(root/ Path('fichiers correctifs') / Path('correction_courriel_secretaire.json'), 'r') as file:
        dic_correction = json.load(file)
    for num_licence in dic_correction.keys():
            idx = df[df["N°"]==int(num_licence)].index
            df.loc[idx,'Adresse email'] = dic_correction[num_licence]
    return df

def add_web(root):
    # Le fichier liste_club_web.xlsx est extrait du site structure FFVelo 
    # Extraction/Liste des structures affililiées
    
    file_mf = root / Path("fichiers FFCT") / Path("liste_club_web.xlsx")
    web_df = pd.read_excel(file_mf)
    dic_web=dict(zip(web_df['Code Structure'],web_df['Adresse Web']))

    return dic_web

def add_seconde_adresse(root):
    
    file_mf = root / Path("fichiers FFCT") / Path("deuxieme_adresse_president.xlsx")
    add_df = pd.read_excel(file_mf)
    dic_add = dict(zip(add_df['N°'],add_df['Adresse électronique']))

    return dic_add

    
def add_feminine(df,nom):
    df = pd.concat([df, pd.DataFrame({"Saison":[2025],
                                      "Groupe d'instance":["Bureau Exécutif"],
                                      "Fonction":["Référent(e) Féminine"],
                                      "Correspondant":["Non"],
                                      "Civilité Nom Prénom":[nom]})], ignore_index=True)
    return df
    
def make_dic_nom_club(root):
    with open(root / Path('nom_club.json'), 'r') as file:
        data = json.load(file)
    return data

def make_dic_ville(root):
    with open(root / Path("ville_club.json"), 'r') as file:
        data = json.load(file)
    return data

def make_dic_territoire(root):
    with open(root / Path("territoire.json"), 'r') as file:
     data = json.load(file)
    return data
    
def make_list_courriel_president(result):
    with open(root / Path('fichiers correctifs') / Path('add_courriel_president.json'), 'r') as file:
        add_list = json.load(file)
    list_president = result.query('P=="P"')['Adresse email'].tolist()+add_list
    return list_president

def merge_row(df):
    
    df = df.fillna('')
    df_list = []
    for x,y in df.groupby('Nom'):
        if len(y)>1:
            for col in ["P","C","S","D","T","J","F"]:
                merge = ''.join(y[col].tolist())
                if len(merge) > 1 : merge = merge[0]
                y[col] = merge
            y = y.drop_duplicates()
        df_list.append(y)
    
    df_tot = pd.concat(df_list, axis=0)
    return df_tot
    
def build_bureau(df_bureau, df_effectif,nom_club_dic,id_club, ville_club_dic, territoire_dic,flag):
    df_merged = df_bureau.merge(df_effectif[['N°','Nom', 'Sexe','Adresse email', 'Tel fixe',
                                             'Tel portable','Adresse','Age' ]], on='Nom', how='left')
    df_merged['P'] = df_merged['Fonction'].apply(lambda x: 'P' if x=='Président(e)' else '')
    df_merged['C'] = df_merged['Correspondant'].apply(lambda x: 'C' if x=='Oui' else '')
    df_merged['S'] = df_merged['Fonction'].apply(lambda x: 'S' if x=='Secrétaire' else '')
    df_merged['D'] = df_merged['Fonction'].apply(lambda x: 'D' if x=='Délégué(e) sécurité' else '')
    df_merged['T'] = df_merged['Fonction'].apply(lambda x: 'T' if x=='Trésorier(e)' else '')
    df_merged['J'] = df_merged['Fonction'].apply(lambda x: 'J' if x=='Référent(e) Jeunes' else '')
    df_merged['F'] = df_merged['Fonction'].apply(lambda x: 'F' if x=='Référent(e) Mixité' else '')
    
    df_merged['ID'] = int(file.split('.')[0])
    df_merged['Nom du Club'] = nom_club_dic[id_club]
    df_merged['Ville du Club'] = ville_club_dic[id_club]
    df_merged['Territoire'] = territoire_dic[id_club]
    df_merged['Flag'] = flag
    df_merged = df_merged[['ID','Nom du Club','Ville du Club','Territoire',
                           'P','C','S','D','T','J','F','N°','Nom','Adresse','Tel fixe',
                           'Tel portable','Adresse email','Age','Flag']]
    df_merged = df_merged.fillna('')
    df_merged['Tel fixe'] = df_merged['Tel fixe'].apply(lambda x: int(x) if x!='' else '')
    df_merged['Tel portable'] = df_merged['Tel portable'].apply(lambda x: int(x) if x!='' else '')
    df_merged = merge_row(df_merged)
    return df_merged

year = 2026
year_effectif = year


root = Path.home() / Path('Dropbox\Codep 38') / Path("Année " + str(year)) / Path('Secretariat')
root_ref = root / Path('FICHIER_REFERENCE')
root_bureau = Path.home() / Path('Dropbox\Codep 38') / Path("Année " + str(year)) 
root_bureau = root_bureau / Path('Secretariat') / Path("bureau_club_38")
root_effectif = Path.home() / Path('Dropbox\Codep 38') / Path("Année " + str(year)) 
root_effectif = root_effectif / Path('Secretariat') / Path('codep_effectif_club')

fonction_list = ['Président(e)','Secrétaire','Délégué(e) sécurité','Trésorier(e)',
                 'Référent(e) Jeunes','Correspondant','Référent(e) Mixité',
                 'Référent(e) Communication','Membre']

animateur_liste = ['Animateur Club', 'Initiateur fédéral', 'Moniteur Fédéral','Instructeur Fédéral']
with open(root / Path ('fichiers correctifs') / Path('referentes_feminines.json'), 'r') as file:
    feminine_dict = json.load(file)

nom_club_dic = make_dic_nom_club(root_ref)
ville_club_dic = make_dic_ville(root_ref)
territoire_dic = make_dic_territoire(root_ref)
bureau_list_df = []
df_animateur_list = []

q = [int(x.split('.')[0]) for x in  os.listdir(root_bureau) if x.endswith('.xlsx')]
q.sort()
files = [str(x)+'.xlsx' for x in q]
flag = 1
flag_animateur = 1
for file in files:
    id_club = file.split('.')[0]
    df_bureau = pd.read_excel(root_bureau / Path(file))
    #if str(id_club) in feminine_dict.keys():
    #    for name in feminine_dict[id_club]:
    #       df_bureau = add_feminine(df_bureau,name)
        
    df_bureau = df_bureau.rename(columns={"Civilité Nom Prénom": "Nom"})
    df_bureau['Nom'] = df_bureau['Nom'].apply(lambda x: ' '.join(x.split()[1:] if isinstance(x,str) else ''))
    df_bureau = df_bureau.fillna('Correspondant')
    df_bureau['Fonction'] = df_bureau['Fonction'].apply(lambda x :  'Président(e)' if 'Co-président(e)' in x else x)
    #df_bureau['Fonction'] = df_bureau['Fonction'].apply(lambda x :  'Président(e)' if 'Vice-Président(e)' in x else x)
    df_bureau = df_bureau.query('Fonction in @fonction_list')

    if "Délégué(e) sécurité" not in df_bureau['Fonction'].tolist():
        ds = df_bureau.query('Fonction=="Président(e)"')
        ds['Fonction'] = "Délégué(e) sécurité"
        df_bureau = pd.concat([df_bureau, ds], axis=0)
        print("sans délégué sécurité",id_club)
    
   
    fil = pd.ExcelFile(root_effectif / Path(file))
    df_effectif = pd.read_excel(root_effectif / Path(file))
    df_effectif['Date de naissance'] = pd.to_datetime(df_effectif['Date de naissance'],
                                                 format="%d/%m/%Y")
    df_effectif['Age']  = df_effectif['Date de naissance'].apply(lambda x :
                                                  (pd.Timestamp(int(year), 9, 30)-x).days/365)
    if int(fil.sheet_names[0].split(' ')[3]) != len(df_effectif):
        print(file)
    df_effectif = df_effectif.rename(columns={"Nom, Prénom": "Nom"})
    df_effectif['Nom'] = df_effectif['Nom'].apply(lambda x: ' '.join(x.split()[1:]))
    df_effectif = df_effectif.rename(columns={"Formation Educateur": "Formation",})
    df_animateur = df_effectif.query('Formation in @ animateur_liste')
    num_club = file.split('.')[0]
    df_animateur['N° Club'] = num_club
    df_animateur['Club'] = nom_club_dic[num_club]
    df_animateur['Flag'] = flag_animateur 
    df_animateur_list.append(df_animateur)
    dg = build_bureau(df_bureau, df_effectif, nom_club_dic, id_club, ville_club_dic, territoire_dic,flag)
    flag += 1 
    
    if len(df_animateur)>0 : flag_animateur  += 1
    bureau_list_df.append(dg)
       
result = pd.concat(bureau_list_df, axis=0 ,ignore_index= True) 
result = correction_email_president(result,root)
result = correction_email_secretaire(result)
result['Tel fixe'] = result['Tel fixe'].apply(normalize_num_tel)
result['Tel portable'] = result['Tel portable'].apply(normalize_num_tel)
site_dic = add_web(root)
result['Site'] = result['ID'].map(site_dic)
add_dic = add_seconde_adresse(root)
result['Adresse email 2'] = result['ID'].map(add_dic)
result['Adresse email 2'] = result.apply(lambda row: row['Adresse email 2'] if row['P']=="P" else "" ,axis=1)
n_president_dic = {}
for x,y in result.groupby(['ID'])['P']:
    n_president_dic[x[0]] = y.tolist().count('P')
result['nbr_president'] = result['ID'].map(n_president_dic)

list_courriel_president = make_list_courriel_president(result)

result = result[['ID', 'Nom du Club', 'Ville du Club', 'Territoire', 'P', 'C', 'S', 'D',
       'T', 'J', 'F', 'N°', 'Nom', 'Adresse', 'Tel fixe', 'Tel portable',
       'Adresse email', 'Adresse email 2','Age','Site', 'nbr_president','Flag']]

root_output = Path.home() / Path('Dropbox\Codep 38') / Path("Année " + str(year))  
result.to_excel(root_output / f'bureau_club_codep38_{date_today()}.xlsx',index=None) 

print(root_output / f'bureau_club_codep38_{date_today