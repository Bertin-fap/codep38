import os
from pathlib import Path

import pandas as pd
import matplotlib.pyplot as plt

def stat_age_codep(year):
    root = Path.home() / Path('Dropbox\Codep 38') / Path("Année " + str(year)) / Path('Secretariat')
    root = root / Path('EVOLUTION') / Path('codep_age')
    
    dic_age = {}
    dic_age_genre = {}
    for file in [x for x in os.listdir(root) if x.endswith('.xlsx')]:
        df = pd.read_excel(root / Path(file),skipfooter=1)
        df = df.fillna(0)
        df = df.fillna(0)
        df['mean']=df.apply(lambda row:row['Âge']*row['Total'],axis=1)
        dic_age[file[0:4]] = [sum(df['mean'])/sum(df['Total'])]
        df['mean_h']=df.apply(lambda row:row['Âge']*row['Homme'],axis=1)
        df['mean_f']=df.apply(lambda row:row['Âge']*row['Femme'],axis=1)
        dic_age_genre[file[0:4]] = [sum(df['mean_h'])/sum(df['Homme']),
                              sum(df['mean_f'])/sum(df['Femme'])]
    
    dg = pd.DataFrame.from_dict(dic_age).T
    dg.columns = ['Age moyen']
    dg.to_excel(r"c:\users\franc\Temp\spy.xlsx")
    ax = dg.plot.bar(rot=90,stacked=True)
    plt.ylim(40,60)
    
    dg = pd.DataFrame.from_dict(dic_age_genre).T
    dg.columns = ['Age moyen homme','Age moyen femme']
    ax = dg.plot.bar(rot=90)
    plt.ylim(40,75)

year = 2026
stat_age_codep(year)