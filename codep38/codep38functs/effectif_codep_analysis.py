import os
from pathlib import Path
import matplotlib.pyplot as plt

import pandas as pd

def stat_effectif_codep(year):

    root = Path.home() / Path('Dropbox\Codep 38') / Path("Année " + str(year)) / Path('Secretariat')
    root = root / Path('EVOLUTION') / Path('codep_effectif')
    
    dic_effectif = {}
    dic_nbr_club = {}
    dic_pourcentage = {}
    for file in [x for x in os.listdir(root) if x.endswith('.xlsx')]:
        df = pd.read_excel(root / Path(file),skipfooter=1,skiprows=2,header=None)
        df.columns = ['id_club','nom_club','renouvellement_homme','renouvellement_femme','creation_homme','creation_femme',
                      'total_homme','total_femme']
        df['total_femme'] =df['total_femme'].fillna(0)
        df['total_homme'] =df['total_homme'].fillna(0)
        dic_effectif[file[0:4]] = [sum(df['total_homme']),sum(df['total_femme'])]
        dic_nbr_club[file[0:4]] = [len(df)]
        dic_pourcentage[file[0:4]] = [100*sum(df['total_femme'])/(sum(df['total_femme'])+sum(df['total_homme']))]
    dg = pd.DataFrame.from_dict(dic_effectif).T
    dg.columns = ['Homme','Femme']
    ax = dg.plot.bar(rot=90,stacked=True)
    plt.grid()
    #print(dg)
    dh = pd.DataFrame.from_dict(dic_nbr_club).T
    dh.columns = ['# club']
    ax = dh.plot.bar(rot=90,stacked=True)
    plt.grid()
    
    dp = pd.DataFrame.from_dict(dic_pourcentage).T
    dp.columns = ['poucentage de femmme']
    ax = dp.plot.bar()

year = 2026
stat_effectif_codep(year)