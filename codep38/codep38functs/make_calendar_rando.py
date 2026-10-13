__all__ ='make_calendar_rando'

from datetime import date
import datetime
import calendar
import re
from collections import defaultdict
from textwrap import wrap
import unicodedata
from tkinter import messagebox

import os
from pathlib import Path

import pandas as pd
import xlrd

from docxtpl import DocxTemplate
from docxcompose.composer import Composer
from docx import Document as Document_compose



from datetime import date
import datetime
import calendar
import re
from collections import defaultdict
from textwrap import wrap
import unicodedata

import os
from pathlib import Path

import pandas as pd
import xlrd

from docxtpl import DocxTemplate
from docxcompose.composer import Composer
from docx import Document as Document_compose



VERSION = '01'
semi_quadratin = "\u2014"
JOURS_DIC = {"Monday" :"lundi",
             "Tuesday" :"mardi",
             "Wednesday" :"mercredi",
             "Thursday" :"jeudi",
             "Friday" :"vendredi",
             "Saturday" :"samedi",
             "Sunday" :"dimanche",}

MONTHS_DIC = {"janvier" : 1,
              "février" : 2,
              "mars" : 3,
              "avril" : 4,
              "mai" : 5,
              "juin" : 6,
              "juillet" : 7,
              "août" : 8,
              "septembre" : 9,
              "octobre" : 10,
              "novembre" : 11,
              "décembre" : 12,}

INV_MONTH_DIC = dict(zip(MONTHS_DIC.values(),MONTHS_DIC.keys()))

templates_dict = {'Route': 'template_randonnees_route.docx',
                  'Route/ Gravel': 'template_randonnees_route_gravel.docx',
                  'Route/ Pédestre': 'template_randonnees_route_pedestre.docx',
                  'Route/ Vtt': 'template_randonnees_route_vtt.docx',
                  'Route/ Vtt/ Gravel': 'template_randonnees_route_vtt_gravel.docx',
                  'Route/ Vtt/ Pédestre': 'template_randonnees_route_vtt_pedestre.docx',
                  "BOURSE AUX VELOS": 'template_randonnees_bourse_aux_velos.docx',
                  'Vtt/ Pédestre': 'template_randonnees_vtt_pedestre.docx',
                  'Vtt':'template_randonnees_vtt.docx',
                  'Vtt/ Gravel':'template_randonnees_vtt_gravel.docx',
                  'Route/ Vtt/ Pédestre/ Gravel': 'template_randonnees_route_vtt_pedestre_gravel.docx',
                  'Route/ Pédestre/ Gravel': 'template_randonnees_route_pedestre_gravel.docx',
                  'Pédestre/ Gravel':'template_randonnees_pedestre_gravel.docx',
                  'Vtt/ Pédestre/ Gravel':'template_randonnees_vtt_pedestre_gravel.docx',
                  'Vtt/ Pédestre':'template_randonnees_vtt_pedestre.docx'}

def make_table_matiere(df,PATH_OUTPUT,PATH_TEMPLATES,year):

    output_file = PATH_OUTPUT / Path( 'liste_matiere.docx')
    pdep = 5
    frameworks = []
    for idx,row in df.iterrows():
        if 'BAC' in row['nom_randonnee']:
            date = '19-20 Juin'
        else:
            date = row['date'].split(' ')[1] + ' ' + row['mois']
        
        frameworks.append(dict(mois = date,
                               typ = row['Type de manifestation'],
                               name = row['nom_randonnee'],
                               page = pdep + idx//2))
                
    context ={'year': year,
              'frameworks': frameworks}
    template_docx = PATH_TEMPLATES  / Path('template_Liste_matiere.docx')
    doc = DocxTemplate(template_docx) 
    doc.render(context)
    doc.save(output_file)
    
def concat_tel(row):
    if len(row["tel_fixe"]) > 0:
        return row["tel_fixe"] + " / " + row["tel_mobile"]
    else:
        return row["tel_mobile"]

def has_numbers(inputString):
    return any(char.isdigit() for char in inputString)

def normalize_euro(text):

    rep = {"euros": "€", "euro": "€", "Euros":"€","Euro":"€"} 
    pattern = re.compile("|".join(rep.keys()))
    text = pattern.sub(lambda m: rep[re.escape(m.group(0))], text)
    return text

def combine_all_docx(list_docx,year,PATH_TEMPLATES,PATH_OUTPUT):

    """
    Merges all the docx file contained in `list_docx` 
    """

    filename_master = PATH_TEMPLATES / Path('premiere-de-couverture.docx')
    number_of_sections = len(list_docx)
    master = Document_compose(filename_master)
    composer = Composer(master)
    for i in range(0, number_of_sections):
        doc_temp = Document_compose(list_docx[i])
        try:
            composer.append(doc_temp)
        except:
            pass
            
    file_name = f'CALENDRIER_{str(year)}.docx'
    combine_file = PATH_OUTPUT / file_name
    print(combine_file)
    composer.save(combine_file)

def builds_organisation_docx(df,year,PATH_TEMPLATES,PATH_OUTPUT):
    df.to_excel(r"C:\users\franc\Temp\spy.xlsx")
    long = 1  # for future use
    file_list = []
    list_docx = []
    idx = 0
    for i_dep in range(0,len(df),long):
        i_fin = i_dep+long
        template_path_docx = df['template'].tolist()[i_dep:i_fin][0]
        template_path_docx = PATH_TEMPLATES / Path(template_path_docx)
        doc = DocxTemplate( template_path_docx)
        
        f ={'version': VERSION,
            'mois': df['mois'].tolist()[i_dep:i_fin],
            'jour': df['date'].tolist()[i_dep:i_fin],
            'titre1': df['nom_randonnee'].tolist()[i_dep:i_dep+long],
            'titre2': df['club_organisateur'].tolist()[i_dep:i_fin],
            'type': df['Type de manifestation'].tolist()[i_dep:i_fin],
            'lieu1': df['adresse1'].tolist()[i_dep:i_fin],
            'lieu2': df['adresse2'].tolist()[i_dep:i_fin],
            'horaire': df['horaire'].tolist()[i_dep:i_fin],
            'kr': df['distances_routes'].tolist()[i_dep:i_fin],
            'kv': df['distances_vtt'].tolist()[i_dep:i_fin],
            'kp': df['distances_pedestre'].tolist()[i_dep:i_fin],
            'kg':  df['distances_gravel'].tolist()[i_dep:i_fin],
            'tffct': df['tarifs_ffct_route'].tolist()[i_dep:i_fin],
            'tffct_18': df['tarifs_ffct_route_18'].tolist()[i_dep:i_fin],
            'aut': df['tarifs_non_ffct_route'].tolist()[i_dep:i_fin],
            'aut_18': df['tarifs_non_ffct_route_18'].tolist()[i_dep:i_fin],
            'tp': df['tarifs_pedestre'].tolist()[i_dep:i_fin],
            'tp_18': df['tarifs_pedestre_18'].tolist()[i_dep:i_fin],
            'nom': df['contact'].tolist()[i_dep:i_fin],
            'tel': df['tel_mobile'].tolist()[i_dep:i_fin],
            'mail': df['courriel'].tolist()[i_dep:i_fin],
            'web': df['web'].tolist()[i_dep:i_fin]}
            
        doc.render(f)
        #    
        output_file = PATH_OUTPUT / Path("5-sejour"+str(idx)+".docx")
        doc.save(output_file)
        list_docx.append(output_file)
    
        idx += 1
        
    return list_docx
       

def normalize_time(s):
    l = re.split(r"[\D]+", str(s).strip())
    minutes = l[1].zfill(2) if len(l) ==2 else '00'
    t = l[0].zfill(2)+ " h " + minutes
    return t

def dic_rando_(year,root):
    
    YEAR = year
    jours = ['lundi','mardi','mercredi','jeudi','vendredi','samedi','dimanche']
    
    
    def normalize_num_tel(s):
        if isinstance(s,float) : s = int(s)
        s = str(s)
        if len(s)>0:
            digits = ''.join(char for char in s if char.isdigit())
            digits = digits.zfill(10)
            return " ".join(wrap(digits, width=2))
        else:
            return ''
        
    def parse_date(worksheet,workbook,num_row):
        my_date = worksheet.cell_value(num_row,1)
    
        if isinstance(my_date,float):
            my_date = datetime.datetime(*xlrd.xldate_as_tuple(my_date, workbook.datemode))
            my_date = my_date.replace(year=year)
        else:
            my_date_split = my_date.split()
            if my_date_split[0].lower() in jours:
                my_date = datetime.datetime(YEAR,MONTHS_DIC[my_date_split[2].lower()], int(my_date_split[1]))
            elif "/" in my_date:
                my_date_split = my_date.split('/')
                my_date = datetime.datetime(YEAR,int(my_date_split[1]), int(my_date_split[0]))
            elif "," in my_date:
                my_date_split = my_date.split(',')
                my_date = datetime.datetime(YEAR,int(my_date_split[1]), int(my_date_split[0]))    
            elif str(YEAR) in my_date:
                print(file,my_date)
                my_date = datetime.datetime(YEAR,int(my_date_split[1]), int(my_date_split[0]))
            else:
                print(f'Cannot parse date in file : {file}')
                my_date = datetime.datetime(YEAR,1, 1)
        return my_date
        
    def parse_nom_randonnee(worksheet,num_row):
        return worksheet.cell_value(num_row,4).strip().upper()
    
    def parse_club_organisateur(worksheet,num_row):
        return worksheet.cell_value(num_row,4).strip()
    
    def parse_numero_federal(worksheet,num_row):
        numero_federal = worksheet.cell_value(num_row,23)
        if numero_federal:
            return str(int(numero_federal)).zfill(5)
        else:
            return ''
    
    def parse_adresse_1(worksheet,num_row):
        return worksheet.cell_value(num_row,3)

    def parse_adresse_2(worksheet,num_row):
        return worksheet.cell_value(num_row,3)
    
    def parse_horaire_ouverture_deb(worksheet,workbook,num_row):
        cell_type = worksheet.cell_type(num_row,4)
        cell_value = worksheet.cell_value(num_row,4)
        if cell_type == xlrd.XL_CELL_DATE:
            dt_tuple = xlrd.xldate_as_tuple(cell_value, workbook.datemode)
            t = str(dt_tuple[3]).zfill(2) + ' h ' + str(dt_tuple[4]).zfill(2)
            return t
        else:
            return normalize_time(worksheet.cell_value(num_row,4))
            
    def parse_horaire_ouverture_fin(worksheet,workbook,num_row):
        cell_type = worksheet.cell_type(num_row,6)
        cell_value = worksheet.cell_value(num_row,6)
        if cell_type == xlrd.XL_CELL_DATE:
            dt_tuple = xlrd.xldate_as_tuple(cell_value, workbook.datemode)
            t = str(dt_tuple[3]).zfill(2)+ ' h ' + str(dt_tuple[4]).zfill(2)
            return t
        else:
            return normalize_time(worksheet.cell_value(num_row,6))  
            
    def parse_horaire_fermeture(worksheet,workbook,num_row):
        cell_type = worksheet.cell_type(num_row,4)
        cell_value = worksheet.cell_value(num_row,4)
        if cell_type == xlrd.XL_CELL_DATE:
            dt_tuple = xlrd.xldate_as_tuple(cell_value, workbook.datemode)
            t = str(dt_tuple[3]).zfill(2)+ ' h ' + str(dt_tuple[4]).zfill(2)
            return t
        else:
            return normalize_time(worksheet.cell_value(num_row,4))

    def parse_distance_route(worksheet,num_row):
        lst = [str(int(x)) for x in [worksheet.cell_value(num_row,5),
                worksheet.cell_value(num_row,6),
                worksheet.cell_value(num_row,7),
                worksheet.cell_value(num_row,8),
                worksheet.cell_value(num_row,9),
                worksheet.cell_value(num_row,10),
                worksheet.cell_value(num_row,11),
                worksheet.cell_value(num_row,12)] if isinstance(x,float)]
        lst.sort()
        distance_route = ', '.join(lst)
        return distance_route
        
    def parse_distance_vtt(worksheet, num_row):
        lst = [str(int(x)) for x in [worksheet.cell_value(num_row,5),
                worksheet.cell_value(num_row,6),
                worksheet.cell_value(num_row,7),
                worksheet.cell_value(num_row,8),
                worksheet.cell_value(num_row,9),
                worksheet.cell_value(num_row,10),
                worksheet.cell_value(num_row,11),
                worksheet.cell_value(num_row,12)] if isinstance(x,float)]
        lst.sort()
        distance_vtt = ', '.join(lst)
        return distance_vtt
        
    def parse_distance_pedestre(worksheet, num_row):
        lst = [str(int(x)) for x in [worksheet.cell_value(num_row,5),
                worksheet.cell_value(num_row,6),
                worksheet.cell_value(num_row,7),
                worksheet.cell_value(num_row,8),
                worksheet.cell_value(num_row,9),
                worksheet.cell_value(num_row,10),
                worksheet.cell_value(num_row,11),
                worksheet.cell_value(num_row,12)] if isinstance(x,float)]
        lst.sort()
        distance_pedestre =  ', ' .join(lst)
        return distance_pedestre

    def parse_distance_gravel(worksheet, num_row):
        lst = [str(int(x)) for x in [worksheet.cell_value(num_row,5),
                worksheet.cell_value(num_row,6),
                worksheet.cell_value(num_row,7),
                worksheet.cell_value(num_row,8),
                worksheet.cell_value(num_row,9),
                worksheet.cell_value(num_row,10),
                worksheet.cell_value(num_row,11),
                worksheet.cell_value(num_row,12)] if isinstance(x,float)]
        lst.sort()
        distance_gravel =  ', ' .join(lst)
        return distance_gravel
    
    def parse_tarifs_ffct(worksheet,num_row):
        
        tarifs_ffc_route = [worksheet.cell_value(num_row,5),
                worksheet.cell_value(num_row,6),
                worksheet.cell_value(num_row,7),
                worksheet.cell_value(num_row,8),
                worksheet.cell_value(num_row,9),
                worksheet.cell_value(num_row,10),
                worksheet.cell_value(num_row,11),
                worksheet.cell_value(num_row,12)]
        return tarifs_ffc_route[0]

    def parse_tarifs_ffct_18(worksheet,num_row):
        x_raw = worksheet.cell_value(num_row,22)
        x = str(x_raw).strip()
        if len(x) == 0:
           return ''
        if '- 18 ans :' in x:
            y = x.split(':')[1]
            if isinstance(y,str):
                if not has_numbers(y):
                    return y.lower()
                
                elif y.strip().endswith('€'):
                    y = y.replace(",", ".") 
                    return str(int(float(y.split('€')[0].strip())))+" €"
                else:
                    if y.strip():
                        return str(int(y))+" €"
                    else:
                        return ''     
        else:
            if isinstance(x_raw,int) or isinstance(x_raw,float):
                return str(int(x_raw))+' €'
            return x
        return ''
        
    def parse_tarifs_non_ffct(worksheet,num_row):
        tarifs_ffc_vtt =  [worksheet.cell_value(num_row,5),
                worksheet.cell_value(num_row,6),
                worksheet.cell_value(num_row,7),
                worksheet.cell_value(num_row,8),
                worksheet.cell_value(num_row,9),
                worksheet.cell_value(num_row,10),
                worksheet.cell_value(num_row,11),
                worksheet.cell_value(num_row,12)]
        return tarifs_ffc_vtt[0]


    def parse_tarifs_non_ffct_18(worksheet,num_row):
        x_raw = worksheet.cell_value(num_row,22)
        x = str(x_raw).strip()
        x = normalize_euro(x)
        if len(x) == 0:
           return ''
        if '- 18 ans non ffct tarif :' in x:
            y = x.split(':')[1]
            if isinstance(y,str):
                if 'gratuit' in y.lower():
                    return 'gratuit'
                elif 'non admis' in y.lower():
                    return 'non admis'
                elif y.strip().endswith('€'):
                    y = y.replace(",", ".") 
                    return str(int(float(y.split('€')[0].strip())))+" €"
                else:
                    if y.strip():
                        return str(int(y))+" €"
                    else:
                        return ''     
        else:
            if isinstance(x_raw,int) or isinstance(x_raw,float):
                return str(int(x_raw))+' €'
            return x
        return ''
        
    def parse_tarifs_pedestre(worksheet,num_row):
        tarifs_ffc_pedestre = [worksheet.cell_value(num_row,5),
                worksheet.cell_value(num_row,6),
                worksheet.cell_value(num_row,7),
                worksheet.cell_value(num_row,8),
                worksheet.cell_value(num_row,9),
                worksheet.cell_value(num_row,10),
                worksheet.cell_value(num_row,11),
                worksheet.cell_value(num_row,12)]
        return tarifs_ffc_pedestre[0]

    def parse_tarifs_pedestre_18(worksheet,num_row):
        tarif = worksheet.cell_value(num_row,22)
        if '- 18 ans :' in tarif:
            return tarif.split(':')[1]
        return ''
        
    def parse_nom(worksheet,num_row):
        nom = (worksheet.cell_value(num_row,1)).upper()
        return nom

    def parse_prenom(worksheet,num_row):
        prenom = (worksheet.cell_value(num_row,5))
        prenom = "-".join([x.capitalize() for x in re.split( '[-\s]',prenom)])
        return prenom

    def parse_adresse_contact(worksheet,num_row):
        return (worksheet.cell_value(num_row,1))

    def parse_code_postal(worksheet,num_row):
        return (worksheet.cell_value(num_row,1))

    def parse_ville(worksheet,num_row):
        return (worksheet.cell_value(num_row,5))

    def parse_tel_fixe(worksheet,num_row):
        tel = worksheet.cell_value(num_row,2)
        tel = normalize_num_tel(tel)
        return tel

    def parse_tel_mobile(worksheet,num_row):
        tel = normalize_num_tel(worksheet.cell_value(num_row,6))
        return tel

    def parse_courriel(worksheet,num_row):
        return (worksheet.cell_value(num_row,1))  

    def parse_web(worksheet,num_row):
        return (worksheet.cell_value(num_row,2))  

    def analyse_worsheet(worksheet):
        item_dic = {}
        j=0
        for i in range(4,41):
            cell = worksheet.cell_value(i,0).strip()
            if cell == "Tarifs":
                item_dic[cell+str(j)] = i
                j += 1
            elif cell != '':
                item_dic[cell] = i
            if  cell =="Site Internet :" :
                break
        return  item_dic
        
    dic_rando = defaultdict(list) 
    
    for file in [x for x in os.listdir(root) if x.endswith('.xls')]:
        
        print(file)
        df = pd.read_excel(Path(root)/ Path(file))
        workbook = xlrd.open_workbook(root / Path(file), formatting_info=True)
        worksheet = workbook.sheet_by_index(0)
        type_rando = []

        item_dic = analyse_worsheet(worksheet)
        
        my_date = parse_date(worksheet,workbook,item_dic['Date:'])
        week_day = JOURS_DIC[calendar.day_name[my_date.weekday()]]
    
        month = INV_MONTH_DIC[my_date.month]

        nom_randonnee = parse_nom_randonnee(worksheet,item_dic['Nom de la randonnée :'])
        club_organisateur = parse_club_organisateur(worksheet,item_dic['Nom du club organisateur :'])
        numero_federal = parse_numero_federal(worksheet,item_dic['Nom du club organisateur :'])
        adresse1 = ''
        adresse2 = ''
        if 'Adresse accueil' in item_dic.keys():
            adresse1 = parse_adresse_1(worksheet,item_dic['Adresse accueil'])
            adresse2 = parse_adresse_2(worksheet,item_dic['Adresse accueil']+1)
        horaire_ouverture_deb = parse_horaire_ouverture_deb(worksheet,workbook,item_dic['Horaires ouvertures'] )
        horaire_ouverture_fin = parse_horaire_ouverture_fin(worksheet,workbook,item_dic['Horaires ouvertures'])
        horaire_fermeture = parse_horaire_fermeture(worksheet,workbook,item_dic['Horaires clôtures'] )


        if 'bourse' in nom_randonnee.lower():
            type_rando.append("BOURSE AUX VELOS")
        else:
            distances_route = ''
            if 'Distances parcours Route' in item_dic.keys():
                distances_route = parse_distance_route(worksheet,item_dic['Distances parcours Route'])
                if len(distances_route)> 0 :type_rando.append( "Route")
                
            distances_vtt = ''
            if 'distances parcours VTT' in item_dic.keys():
                distances_vtt = parse_distance_vtt(worksheet,item_dic['distances parcours VTT'])
                if len(distances_vtt)> 0 :type_rando.append("Vtt")
    
            distances_pedestre = ''
            if 'Distances  parcours Pedestre' in item_dic.keys():
                distances_pedestre = parse_distance_pedestre(worksheet,item_dic['Distances  parcours Pedestre'])
                if len(distances_pedestre)> 0 : type_rando.append("Pédestre")
    
            distances_gravel = ''
            if 'Distances parcours Gravel' in item_dic.keys():
                distances_gravel = parse_distance_gravel(worksheet,item_dic['Distances parcours Gravel'])
                if len(distances_gravel)> 0 : type_rando.append("Gravel")
      
        type_rando = '/ '.join(type_rando)
            
        tarifs_ffct = parse_tarifs_ffct(worksheet,item_dic['Tarifs0'])
        tarifs_ffct_18 =  parse_tarifs_ffct_18(worksheet,item_dic['Tarifs0'])
        tarifs_non_ffct = parse_tarifs_non_ffct(worksheet,item_dic['Tarifs1'])
        tarifs_non_ffct_18 = parse_tarifs_non_ffct_18(worksheet,item_dic['Tarifs1'])
        
        
        tarifs_pedestre = ''
        tarifs_pedestre_18 = ''
        if 'Tarifs2' in item_dic.keys():
            tarifs_pedestre = parse_tarifs_pedestre(worksheet,item_dic['Tarifs2'])
            tarifs_pedestre_18 = parse_tarifs_pedestre_18(worksheet,item_dic['Tarifs2'])
            
        nom_contact = parse_nom(worksheet,item_dic['Nom :'])
        prenom_contact = parse_prenom(worksheet,item_dic['Nom :'])
        contact = f'{prenom_contact} {nom_contact}'
        adresse_contact = parse_adresse_contact(worksheet,item_dic['Adresse'] )
        code_postal = parse_code_postal(worksheet,item_dic['Code postal'])
        ville = parse_ville(worksheet,item_dic['Code postal'])
        tel_fixe = parse_tel_fixe(worksheet,item_dic['N° Tel Fixe :'] )
        tel_mob = parse_tel_mobile(worksheet,item_dic['N° Tel Fixe :'])
        courriel = parse_courriel(worksheet,item_dic['E-mail'])
        web = parse_web(worksheet,item_dic['Site Internet :'])

        
        dic_rando['nom_randonnee'].append(nom_randonnee)
        dic_rando['Type de manifestation'].append(type_rando)
        
        if 'BAC' in nom_randonnee:
            dic_rando['date'].append(f'Samedi 19 juin et dimanche 20 juin')
        else:
            dic_rando['date'].append(f'{week_day.capitalize()} {str(my_date.day)}')
            
        dic_rando['mois'].append(str(my_date.month))
        dic_rando['club_organisateur'].append(club_organisateur) 
        dic_rando['numero_federal'].append(numero_federal) 
        dic_rando['adresse1'].append(adresse1)
        dic_rando['adresse2'].append(adresse2)
        dic_rando['horaire_ouverture_deb'].append(horaire_ouverture_deb)
        dic_rando['horaire_ouverture_fin'].append(horaire_ouverture_fin)
        dic_rando['horaire_fermeture'].append(horaire_fermeture)
        dic_rando['distances_routes'].append(distances_route)
        dic_rando['distances_vtt'].append(distances_vtt)
        dic_rando['distances_pedestre'].append(distances_pedestre)
        dic_rando['distances_gravel'].append(distances_gravel)
        dic_rando['tarifs_ffct_route'].append(tarifs_ffct)
        dic_rando['tarifs_ffct_route_18'].append(tarifs_ffct_18)
        dic_rando['tarifs_non_ffct_route'].append(tarifs_non_ffct)
        dic_rando['tarifs_non_ffct_route_18'].append(tarifs_non_ffct_18)
        dic_rando['tarifs_pedestre'].append(tarifs_pedestre)
        dic_rando['tarifs_pedestre_18'].append(tarifs_pedestre_18)
        dic_rando['contact'].append(contact)
        dic_rando['adresse_contact'].append(adresse_contact)
        dic_rando['code_postal'].append(code_postal)
        dic_rando['ville'].append(ville)
        dic_rando['tel_fixe'].append(tel_fixe)
        dic_rando['tel_mobile'].append(tel_mob)
        dic_rando['courriel'].append(courriel)
        dic_rando['web'].append(web)
        dic_rando['fichier'].append(file)
        dic_rando['full_date'].append(my_date)
        
    return dic_rando
    
def make_calendar_rando(codep_path,year):
    PATH_TEMPLATES = Path(codep_path) / Path(r"Reference files\TEMPLATES_CALENDRIER")
    #PATH_OUTPUT = Path.home() / Path('CODEP38') / Path(str(year)) / Path('Calendrier')
    root = Path(codep_path) / Path(f"Année {year}") / Path('Calendrier')
    PATH_OUTPUT = root / Path('TEMP')
    PATH_OUTPUT.mkdir(parents=True, exist_ok=True)
    root = root / Path(r"02 Fiches calendrier 2027 Ok à inscrire dans calendrier")
    
    dic_rando = dic_rando_(year,root)
    
    df = pd.DataFrame(dic_rando)
    df = df.sort_values(by='full_date',ascending=True)
    df["club_organisateur"] = df["club_organisateur"] + \
                                 f" {semi_quadratin} " + df["numero_federal"].astype(str)
    df["horaire"] = "De " + df["horaire_ouverture_deb"] + " à " \
                    + df["horaire_ouverture_fin"] + f" {semi_quadratin} Clôture à " \
                    + df["horaire_fermeture"]
    df["tel_mobile"] = df.apply(concat_tel,axis=1)
    df["tarifs_ffct_route"] = df["tarifs_ffct_route"].apply(lambda x: str(int(x))+ "€" if isinstance(x,float) else "")
    df["tarifs_pedestre"] = df["tarifs_pedestre"].apply(lambda x: str(int(x))+ "€" if isinstance(x,float) else "")
    df["tarifs_non_ffct_route"] = df["tarifs_non_ffct_route"].apply(lambda x: str(int(x))+ "€" if isinstance(x,float) else "")
    month_dic_inv = dict(zip([str(x) for x in MONTHS_DIC.values()],[x.capitalize() for x in MONTHS_DIC.keys()]))
    df['mois'] = df['mois'].map(month_dic_inv )
    df['template'] = df['Type de manifestation'].map(templates_dict)
    
    #df = df.drop(['full_date'],axis=1)
    
    list_docx = builds_organisation_docx(df,year,PATH_TEMPLATES,PATH_OUTPUT)
    make_table_matiere(df,PATH_OUTPUT,PATH_TEMPLATES,year)
    list_docx.insert(0, PATH_OUTPUT / Path('liste_matiere.docx'))
    #list_docx.insert(1, PATH_TEMPLATES / Path('reunion assurance.docx'))

    output_file = root / Path('randos_synthese'+str(year)+'.xlsx')
    df.to_excel(output_file,index = False)
    
    combine_all_docx(list_docx,year,PATH_TEMPLATES,PATH_OUTPUT)
    
    dg = df.groupby(['Type de manifestation'])['nom_randonnee'].count()
    #print(dg)

    messagebox.showinfo("Information", f"votre fichier a été créé sous: {output_file}")