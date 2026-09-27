import unicodedata

import os
import time
from pathlib import Path
from datetime import date
from collections import defaultdict

import pandas as pd

from docxtpl import DocxTemplate, InlineImage, RichText
from docxcompose.composer import Composer
from docx import Document as Document_compose
from docx.shared import Cm, Emu, Inches, Mm

def read_formateur_codep38(year):

    # Le fichier formateur s.csv est disponible sur l'espace structure FFVelo Formations/Extractions
    
    root = Path.home() / Path(r'Dropbox/Codep 38') / Path(f'Année {str(year)}') / Path(r'Secretariat/codep_effectif_club')
    file = 'formateurs.csv'
    df = pd.read_csv(root / Path(file),
                     sep=';',
                     encoding='latin-1')
    return df
    
def build_list_animateurs(df,id_):
    dg = df.query('`Code Structure`==@id_ and `Niveau 2`== "Animateur Club"')
    animateur = ', '.join(dg["Nom & Prénom"])
    dg = df.query('`Code Structure`==@id_ and `Niveau 2`== "Initiateur"')
    initiateur = ', '.join(dg["Nom & Prénom"])
    dg = df.query('`Code Structure`==@id_ and `Niveau 2`== "Moniteur fédéral"')
    moniteur = ', '.join(dg["Nom & Prénom"])
    return animateur,initiateur,moniteur

def date_today():
    today = date.today()
    d = str(today.year)+str(today.month).zfill(2)+str(today.day).zfill(2)
    return d

def make_dic_rando(PATH_RANDO):

    df_rando = pd.read_excel(PATH_RANDO)
    dg = df_rando.groupby(['numero_federal'])
    rando_dict = defaultdict(lambda: '')
    for idx, v in dg:
        rando_dict[idx[0]] = ' ; '.join(v['nom_randonnee'].tolist())
    return rando_dict
    
def combine_all_docx(list_docx,year):

    """
    Merges all the docx file contained in `list_docx` 
    """
    filename_master = PATH_OUTPUT / Path('premiere-de-couverture-bureau.docx')
    number_of_sections = len(list_docx)
    master = Document_compose(filename_master)
    composer = Composer(master)
    for i in range(1, number_of_sections):
        doc_temp = Document_compose(list_docx[i])
        try:
            composer.append(doc_temp)
            time.sleep(0.1)
        except:
            pass
            
    file_name = f'ANNUAIRE_CoDep38_{date_today()}.docx'
    combine_file = PATH_SAUVEGARDE / file_name
    composer.save(combine_file)

def make_page_de_garde(dic_stat,year):
    #qrc_sejour.append(InlineImage(doc,qrc_name,Cm(1.5)))
    output_file = PATH_OUTPUT / Path( 'premiere-de-couverture-bureau.docx')
    
    template_docx = PATH_TEMPLATES  / Path('premiere-de-couverture-bureau-template.docx')
    doc = DocxTemplate(template_docx)
    file_image = PATH_OUTPUT / Path('histo.jpeg')
    
    dic_stat['image1'] = InlineImage(doc,r'C:/Users/franc/CODEP38/Temp/histo.jpeg',Cm(15)) 
    dic_stat['image2'] = InlineImage(doc,r"C:/Users/franc/CODEP38/Temp/nbr_adherents_club38.png",Cm(16))
    dic_stat['year'] = year
    context = dic_stat
    doc.render(context)
    doc.save(output_file)

def make_table_matiere(club_list):

    output_file = PATH_OUTPUT / Path( 'liste_matiere.docx')
    page_dep = 11
    frameworks = []
    for idx,club in enumerate(club_list):
        frameworks.append(dict(club = club,
                              page = page_dep + idx))

    context ={'year': year,
              'frameworks': frameworks}
    template_docx = PATH_TEMPLATES  / Path('template_Liste_matiere_bureau.docx')
    doc = DocxTemplate(template_docx) 
    doc.render(context)
    doc.save(output_file)

def telephonne(x):
    if len(x):
        return x[0]
    return ''

def join_telephonne(t1,t2):
     if t1 == t2:
         t2 = ''
     if len(t1) and len(t2):
         return t1+' / '+t2
     elif len(t1):
         return t1
     else:
         return t2
         
def make_hyperlink(courriel,doc):
    
    rt = RichText()
    rt.add(courriel, 
           url_id=doc.build_url_id(courriel),
           font="Calibri",
           color='#3366CC',
           size = 20)
    
    return rt

def builds_organisation_docx(result,year,template_path_docx,
                             PATH_RANDO,
                             PATH_OUTPUT):

    def correspondante_feminine(id):
        try:
           t1 = telephonne(result.query('ID == @id and F == "F"')['N° Portable'].tolist())
           t2 = telephonne(result.query('ID == @id and F == "F"')['N° Tél'].tolist())
           telephone = join_telephonne(t1,t2)
           return (result.query('ID == @id and F == "F"')['Nom'].tolist()[0],
                   result.query('ID == @id and F == "F"')['Ville'].tolist()[0],
                   telephone,
                   make_hyperlink(result.query('ID == @id and F == "F"')['E-mail'].tolist()[0],doc),)
        except:
            return ('','','','')

    def correspondante(id):
        t1 = telephonne(result.query('ID == @id and C == "C"')['N° Portable'].tolist())
        t2 = telephonne(result.query('ID == @id and C == "C"')['N° Tél'].tolist())
        telephone = join_telephonne(t1,t2)
        try:
           return (result.query('ID == @id and C == "C"')['Nom'].tolist()[0],
                   result.query('ID == @id and C == "C"')['Ville'].tolist()[0],
                   telephone,
                   make_hyperlink(result.query('ID == @id and C == "C"')['E-mail'].tolist()[0],doc),)
        except:
            return ('','','','')

    def correspondante_securite(id):
        t1 = telephonne(result.query('ID == @id and D == "D"')['N° Portable'].tolist())
        t2 = telephonne(result.query('ID == @id and D == "D"')['N° Tél'].tolist())
        telephone = telephone = join_telephonne(t1,t2)
        try:
           return (result.query('ID == @id and D == "D"')['Nom'].tolist()[0],
                               result.query('ID == @id and D == "D"')['Ville'].tolist()[0],
                               telephone,
                               make_hyperlink(result.query('ID == @id and D == "D"')['E-mail'].tolist()[0],doc),)
        except:
            return ('','','','')
    #file_mf = r"C:\Users\franc\Dropbox\Codep 38\Année 2026\Clubs_38_au_12_01_2026.xlsx"
    #web_df = pd.read_excel(file_mf)
    #web_df.rename(columns={'N°': 'ID',}, inplace=True)

    def correspondante_jeune(id):
        t1 = telephonne(result.query('ID == @id and J == "J"')['N° Portable'].tolist())
        t2 = telephonne(result.query('ID == @id and J == "J"')['N° Tél'].tolist())
        telephone = telephone = join_telephonne(t1,t2)
        try:
           return (result.query('ID == @id and J == "J"')['Nom'].tolist()[0],
                               result.query('ID == @id and J == "J"')['Ville'].tolist()[0],
                               telephone,
                               make_hyperlink(result.query('ID == @id and J == "J"')['E-mail'].tolist()[0],doc),)
        except:
            return ('','','','')

    def president(result,id):
        
        l = result.query('ID == @id and P=="P"')
        np = l['Nom'].tolist()[0]
        ap = l['Ville'].tolist()[0]
        t1 = telephonne([l['N° Portable'].tolist()[0]])
        t2 = telephonne([l['N° Tél'].tolist()[0]])
        tp = join_telephonne(t1,t2)
        wp = make_hyperlink(l['E-mail'].tolist()[0],doc)
        nbr_p = 1
        
        add_2 = l['Adresse email 2'].tolist()[0]
        if add_2:
            add_2 = make_hyperlink(add_2,doc)
        else:
            add2 = ''
        np2 = ''
        ap2 = ''
        tp2 = ''
        wp2 = ''
        add_22= ''   
        if len(l)>1:
            nbr_p = 2
            np2 = l['Nom'].tolist()[1]
            ap2 = l['Ville'].tolist()[1]
            t1 = telephonne([l['N° Portable'].tolist()[1]])
            t2 = telephonne([l['N° Tél'].tolist()[1]])
            tp2 = join_telephonne(t1,t2)
            wp2 = make_hyperlink(l['E-mail'].tolist()[1],doc)
            
            add_2 = l['Adresse email 2'].tolist()[1]
            if add_2:
                add_22 = make_hyperlink(add_2,doc)
            else:
                add22 = ''
        return (np,ap,tp,wp,add_2,np2,ap2,tp2,wp2,add_22,nbr_p)         
        
        
    long = 1  # for future use
    file_list = []
    list_docx = []
    club_list = []
    idx = 0
    result.rename(columns={'Nom du Club': 'club', 
                           'Ville du Club': 'ville'}, inplace=True) 
    result = result.fillna(' ')
    df_moniteur = read_formateur_codep38(year)
    for id in result['ID'].unique():
        
        rando_dict = make_dic_rando(PATH_RANDO)
        
        animateur,initiateur,moniteur = build_list_animateurs(df_moniteur,id)
        
        if result.query('ID==@id')['nbr_president'].tolist()[0]==1:
            doc = DocxTemplate(template_path_docx)
        else:
            doc = DocxTemplate(template_path_docx_2)

        club_list.append(str(id).zfill(5) + ' - ' + result.query('ID == @id')['club'].tolist()[0])

        
        web_list = result.query('ID==@id')['Site']
        web_list = [x for x in  web_list if isinstance(x,str)]
        if len(web_list)>0:
            web_site = make_hyperlink(web_list[0],doc)
        else:
            web_site = ''

        (np,ap,tp,wp,add_2,np2,ap2,tp2,wp2,add_22,nbr_p) = president(result,id)
        (nf,af,tf,wf)=correspondante_feminine(id)
        (nd,ad,td,wd)=correspondante_securite(id)
        (nc,ac,tc,wc)=correspondante(id)
        (nj,aj,tj,wj)=correspondante_jeune(id)
        
        secretaire = result.query('ID == @id and S == "S"')['Nom']
        ns_ = ''
        as_ = ''
        ws_ = ''
        ts = ''
        if len(secretaire) >0:
            ns_ = secretaire.tolist()[0]
            as_ = result.query('ID == @id and S == "S"')['Ville'].tolist()[0]
            ws_ = make_hyperlink(result.query('ID == @id and S == "S"')['E-mail'].tolist()[0],doc)
            t1 = telephonne(result.query('ID == @id and S == "S"')['N° Portable'].tolist())
            t2 = telephonne(result.query('ID == @id and S == "S"')['N° Tél'].tolist())
            ts = join_telephonne(t1,t2)

        tresorier = result.query('ID == @id and T == "T"')['Nom']
        nt_ = ''
        at_ = ''
        wt_ = ''
        tt = ''
        if len(tresorier) >0:
            nt_ = tresorier.tolist()[0]
            at_ = result.query('ID == @id and T == "T"')['Ville'].tolist()[0]
            wt_ = make_hyperlink(result.query('ID == @id and T == "T"')['E-mail'].tolist()[0],doc)
            t1 = telephonne(result.query('ID == @id and T == "T"')['N° Portable'].tolist())
            t2 = telephonne(result.query('ID == @id and T == "T"')['N° Tél'].tolist())
            tt = join_telephonne(t1,t2) 
       
        stat = df_merged.query('ID==@id')
        nbr_tot = stat['Nombre total'].tolist()[0]
        nbr_hommes = stat["Nombre d'hommes"].tolist()[0]
        nbr_jeunes = nbr_tot - stat["Nombre d'hommes (>=18 ans)"].tolist()[0] - stat["Nombre dde femmes (>=18 ans)"].tolist()[0]
        nbr_femmes = stat['Nombre de femmes'].tolist()[0]
        age_moyen = stat['Age moyen'].tolist()[0]
        efv = stat['efv'].tolist()[0]
        
        stat_str = f"Effectif :{nbr_tot} ({nbr_femmes} femmes ,"
        stat_str = stat_str + f"{nbr_hommes} hommes) Nombre de jeunes :{nbr_jeunes}  Age moyen : {age_moyen} ans"
        stat_str = stat_str + f"   EFV :{efv.upper()}"
        
        flag_adress = False
        f ={'Nom_club':  result.query('ID == @id')['club'].tolist()[0],
            'ville':  result.query('ID == @id')['ville'].tolist()[0],
            'site':  web_site,
            'id': str(id).zfill(5),
            'np': np,
            'ap': ap if flag_adress else '',
            'tp': tp,
            'wp': wp,
            'w2p': add_2,
            'np2': np2,
            'ap2': ap2 if flag_adress else '',
            'tp2': tp2,
            'wp2': wp2,
            'w2p2': add_22,
            'ns': ns_,
            'as': as_ if flag_adress else '',
            'ts': ts,
            'ws': ws_,
            'nt': nt_,
            'at': at_ if flag_adress else '',
            'tt': tt,
            'wt': wt_,
            'nc': nc,
            'ac': ac if flag_adress else '',
            'tc': tc,
            'wc': wc,
            'nd': nd,
            'ad': ad if flag_adress else '',
            'td': td,
            'wd': wd,
            'nf': nf,
            'af': af if flag_adress else '',
            'tf': tf,
            'wf': wf,
            'nj': nj,
            'aj': aj if flag_adress else '',
            'tj': tj,
            'wj': wj,
            'aclub': animateur,
            'ifed': initiateur,
            'mfed': moniteur,
            'stat': stat_str,
            'rando' : rando_dict[id]}  

        doc.render(f)
        #    
        output_file = PATH_OUTPUT / Path("bureau"+str(idx)+".docx")
        doc.save(output_file)
        list_docx.append(output_file)
    
        idx += 1
        
    return list_docx, club_list

def make_calendar_word(year,result):
    PATH_OUTPUT = Path.home() / Path("CODEP38") / Path('Temp')
    PATH_SAUVEGARDE = Path.home() / Path('Dropbox\Codep 38') / Path("Année " + str(year)) 
    PATH_TEMPLATES = Path.home() / Path('Dropbox\Codep 38') / Path("Année " + str(year)) 
    PATH_TEMPLATES = PATH_TEMPLATES / Path('Secretariat') /  Path("TEMPLATES")
    
    PATH_RANDO = Path.home() / Path('Dropbox\Codep 38') / Path("Année " + str(year)) 
    PATH_RANDO = PATH_RANDO / Path(r"Calendrier\randonnees\SYNTHESE") / Path(f"randos_synthese{year}_save.xlsx")
    
    template_path_docx = PATH_TEMPLATES / Path("Template_bureau.docx")
    template_path_docx_2 = PATH_TEMPLATES / Path("Template_bureau_2.docx")
    #df_tot_animateurs.rename(columns={'N° Club':'IDC'},inplace=True)
    list_docx,  club_list = builds_organisation_docx(result,year,template_path_docx,PATH_RANDO,PATH_OUTPUT)
    make_table_matiere(club_list)
    list_docx.insert(0, PATH_OUTPUT / Path('liste_matiere.docx'))
    make_page_de_garde(dic_stat,year)
    list_docx.insert(0, PATH_OUTPUT / Path('premiere-de-couverture-bureau.docx'))
    combine_all_docx(list_docx,year)