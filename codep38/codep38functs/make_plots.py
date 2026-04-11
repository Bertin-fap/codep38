import numpy as np
def addlabels(x,y,lab):
    for i in range(len(x)):
        plt.text(x[i]-0.15,y[i]-0.5,lab[i],size=7)
        
df_merged = df_merged.sort_values(['Nombre total'])
df = df_merged[['Nom du Club','Nombre de femmes',"Nombre d'hommes"]]
#df.index = df_merged['Nom du Club']
ax = df.plot(kind='barh', stacked=True, figsize=(6,8))
ax.tick_params(axis='y', labelsize=7) 
lab = [str(round(100*x/(x+y),1))+"%" for x,y in zip(df['Nombre de femmes'],df["Nombre d'hommes"])]
x = [x+y for x,y in zip(df['Nombre de femmes'],df["Nombre d'hommes"])]
addlabels(x,
          range(len(df)),lab)
plt.yticks(range(len(df)), df['Nom du Club'])
plt.grid(which='major', axis='x')
plt.grid(which='minor', linestyle='--',axis='x')


df['tot']=df['Nombre de femmes']+df["Nombre d'hommes"]
df = df.sort_values(['tot'],ascending=False)
y = 100*np.cumsum(df['tot'].tolist())/df['tot'].sum()
x = list(range(1,len(y)+1))
plt.plot(y[::-1],x)
#plt.xlabel('Nombre de club')
#plt.ylabel('% du contingent')
#plt.grid(which='major', axis='x')
plt.grid(which='minor', linestyle='--',axis='x')
np.where(y>50)[0][0]
#plt.title('Sales Distribution by Product Type and Region (Horizontal)')
#plt.xlabel('Sales')
#plt.ylabel('Product Type')
plt.tight_layout()
file = PATH_OUTPUT / Path('nbr_adherents_club38.png')
plt.savefig(file)