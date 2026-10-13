import os
from pathlib import Path
import matplotlib.pyplot as plt

import pandas as pd

import os
from pathlib import Path

import pandas as pd
import matplotlib.pyplot as plt

def stat_age_codep38():
    root = Path.home() / Path('Dropbox\Codep 38')  / Path('Reference files') /Path(r"EVOLUTION\codep38_licences")
    
    dic_age_genre = {}
    dic_age = {}
    for file in [x for x in os.listdir(root) if x.endswith('.csv')]:
        year = int(file[0:4])
        df = df = pd.read_csv(root / Path(file),
                     sep=';',
                     encoding='latin-1')
        df['Date de naissance'] = pd.to_datetime(df['Date de naissance'],
                                                 format="%d/%m/%Y")
        df['Age']  = df['Date de naissance'].apply(lambda x :
                                                  (pd.Timestamp(int(year), 9, 30)-x).days)/365
        
        df['Age'] = df['Age'].astype(float).apply(lambda x : round(x,2))
        
        dic_age[year] = [mean(df['Age']),
                         mean(df.query('Age>18')['Age'])]
        dic_age_genre[year] = [mean(df.query('Sexe=="M"')['Age']),
                               mean(df.query('Sexe=="F"')['Age'])]
    
    dg = pd.DataFrame.from_dict(dic_age).T
    dg.columns = ['Age moyen total','Age moyen adulte']
    dg.to_excel(r"c:\users\franc\Temp\spy.xlsx")
    ax = dg.plot.bar(rot=90)
    plt.ylim(40,70)
    
    dg = pd.DataFrame.from_dict(dic_age_genre).T
    dg.columns = ['Age moyen homme','Age moyen femme']
    ax = dg.plot.bar(rot=90)
    plt.ylim(40,75)

def stat_effectif_codep38():

    root = Path.home() / Path('Dropbox\Codep 38')  / Path('Reference files') /Path(r"EVOLUTION\codep38_licences")
    
    dic_effectif = {}
    dic_nbr_club = {}
    dic_pourcentage = {}
    for file in [x for x in os.listdir(root) if x.endswith('.csv')]:
        year = int(file[0:4])
        df = df = pd.read_csv(root / Path(file),
                     sep=';',
                     encoding='latin-1')
        df['Date de naissance'] = pd.to_datetime(df['Date de naissance'],
                                                 format="%d/%m/%Y")
        df['Age']  = df['Date de naissance'].apply(lambda x :
                                                  (pd.Timestamp(int(year), 9, 30)-x).days)/365
        
        df['Age'] = df['Age'].astype(float).apply(lambda x : round(x,2))
        dfm = df.query('Sexe=="M"')
        dff = df.query('Sexe=="F"')
        nbr_hommes = len(dfm)
        nbr_femmes = len(dff)
            
        dic_effectif[file[0:4]] = [nbr_hommes,nbr_femmes]
        dh = df.groupby(["Nom Structure"])
        dic_nbr_club[year] = [len(dh)]
        dic_pourcentage[year] = [100*nbr_femmes/(nbr_femmes+nbr_hommes)]
    dg = pd.DataFrame.from_dict(dic_effectif).T
    dg.columns = ['Homme','Femme']
    ax = dg.plot.bar(rot=90,stacked=True)
    plt.grid()
    plt.show()
    dh = pd.DataFrame.from_dict(dic_nbr_club).T
    dh.columns = ['# club']
    ax = dh.plot.bar(rot=90,stacked=True)
    plt.grid()
    plt.show()
    
    dp = pd.DataFrame.from_dict(dic_pourcentage).T
    dp.columns = ['poucentage de femmme']
    ax = dp.plot.bar()
    plt.show()
