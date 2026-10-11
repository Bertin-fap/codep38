__all__ ='plot_stat_rando'

from collections import defaultdict
import os
import math
import pathlib
from pathlib import Path
from operator import itemgetter

# Third party imports
import xlrd
import pandas as pd
import matplotlib.pyplot as plt

def addlabels(x,y,lab):
    for i in range(len(x)):
        plt.text(x[i]-0.15,y[i],lab[i],size=7)

def nan2zero(s):
    if isinstance(s,float):
        return int(s)
    return 0
    
def extraction(root):
    d = defaultdict(list)
    club_route = defaultdict(list)
    data = defaultdict(list)
    #club_gravel = defaultdict(list)
    for file in [x for x in os.listdir(root) if x.endswith('.xls') ]:
        workbook = xlrd.open_workbook(Path(root) / Path(file), formatting_info=True)
        worksheet_route = workbook.sheet_by_name('1 ROUTE ')
        worksheet_vtt = workbook.sheet_by_name('2 VTT')
        worksheet_pedestre = workbook.sheet_by_name('4 MARCHE')
        worksheet_gravel = workbook.sheet_by_name('3 GRAVEL')
        if worksheet_route.cell_value(3,2) != '':
            nom_rando = worksheet_route.cell_value(3,2)
            d['nom_rando'].append(nom_rando )
        else:
            nom_rando = worksheet_vtt.cell_value(3,2)
            d['nom_rando'].append(nom_rando)
        d['route'].append(worksheet_route.cell_value(16,7))
        d['vtt'].append(worksheet_vtt.cell_value(16,7))
        d['pedestre'].append(worksheet_pedestre.cell_value(16,7))
        d['gravel'].append(worksheet_gravel.cell_value(16,7))
        d['non_ffct'].append(worksheet_pedestre.cell_value(15,7)+
                             worksheet_route.cell_value(15,7)+
                             worksheet_gravel.cell_value(15,7)+
                             worksheet_vtt.cell_value(15,7))
        i = 21
        while worksheet_route.cell_value(i,1) or i>100:
            data['rando'].append(nom_rando)
            data['type'].append('route')
            data['club'].append(worksheet_route.cell_value(i,2))
            data['dep'].append(str(int(worksheet_route.cell_value(i,3))))
            data['nbr_cyclos'].append(worksheet_route.cell_value(i,7))

            club_route[nom_rando].append((worksheet_route.cell_value(i,1),
                                    worksheet_route.cell_value(i,7),'route'))
            i += 1
        
        i = 21
        while worksheet_vtt.cell_value(i,1) or i>100:
            data['rando'].append(nom_rando)
            data['type'].append('vtt')
            data['club'].append(worksheet_vtt.cell_value(i,2))
            data['dep'].append(str(int(worksheet_vtt.cell_value(i,3))))
            data['nbr_cyclos'].append(worksheet_vtt.cell_value(i,7))

            club_route[nom_rando].append((worksheet_vtt.cell_value(i,1),
                                    worksheet_route.cell_value(i,7),'vtt'))
            i += 1
        
        i = 21
        while worksheet_gravel.cell_value(i,1) or i>100:
            data['rando'].append(nom_rando)
            data['type'].append('gravel')
            data['club'].append(worksheet_gravel.cell_value(i,2))
            data['dep'].append(str(int(worksheet_gravel.cell_value(i,3))))
            data['nbr_cyclos'].append(worksheet_gravel.cell_value(i,7))
            club_route[nom_rando].append((worksheet_gravel.cell_value(i,1),
                                          
                                    worksheet_route.cell_value(i,7),'gravel'))
            i += 1  
        d['# cub_ffct'].append(len(set(map(itemgetter(0), club_route[nom_rando]))))
    return d,club_route,data

def plot_stat_rando(codep_path,year):

    root = Path(codep_path) / Path(f'Année {str(year)}') / Path(r"Calendrier/fiches challenge")
    d,club_route,data = extraction(root)
    
    df = pd.DataFrame(d)
    df['tot'] = df['route'] + df['vtt'] + df['pedestre']+df['gravel']
    df = df.sort_values(['tot'])
    df.to_excel(root / Path('STAT') / Path('participation_rando_club38.xlsx'),index=False)
    #df = df.drop('tot', axis=1)
    df.index = df['nom_rando']
    
    ax = df[['route','vtt','pedestre','gravel']].plot(kind='barh', stacked=True, figsize=(8,8))
    ax.tick_params(axis='y', labelsize=7)
    x = [int(x+y+z+v) for x,y,z,v in zip(df['route'],
                                         df["vtt"],
                                         df["pedestre"],
                                         df["gravel"]) ]
    lab = [str(round(100-100*int(w)/int(x+y+z+v),1))+"%" if int(x+y+z+v)>0 else '' for x,y,z,v,w in zip(df['route'],
                                         df["vtt"],
                                         df["pedestre"],
                                         df["gravel"],
                                         df["non_ffct"]) ]
    addlabels(x,
              range(len(df)),lab)
    plt.yticks(range(len(df)), df['nom_rando'])
    
    plt.ylabel('')
    plt.grid(which='major', axis='x')
    
    
    plt.tight_layout()
    plt.savefig(root / Path('STAT') /Path('participation_rando_club38.png'))
    plt.show()
    
    #dg = pd.DataFrame.from_dict(data)
    #dg.to_excel(r'c:\users\franc\Temp\spy.xlsx',index
