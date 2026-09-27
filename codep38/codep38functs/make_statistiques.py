def stat_calendar(result):
    PATH_OUTPUT = Path.home() / Path("CODEP38") / Path('Temp')
    
    indicator_df = result[['ID','Nom du Club','Ville du Club','Territoire']].drop_duplicates()
    
    dic_effectif,effectif_list_h, effectif_list_f = make_dic_age(root,year)
    effectif_df = pd.DataFrame.from_dict(dic_effectif).T
    effectif_df.columns=["ID","Age moyen","Age moyen homme","Age moyen femme","Age moyen adulte (>=18 ans)",
                         "Age moyen adulte (<18 ans)",
                         "Nombre d'hommes", "Nombre d'hommes (>=18 ans)",
                         "Nombre de femmes", "Nombre dde femmes (>=18 ans)",
                         "Pourcentage femme","Nombre total","efv"]
    df_merged = indicator_df.merge(effectif_df, on='ID', how='left')
    df_merged["points_AG"] = df_merged["Nombre total"].apply(nbr_point_vote)
    df_merged.to_excel(root / Path('indicateurs_club.xlsx'),index=None)
    age_moyen_copep38 =round( sum(effectif_df["Age moyen homme"]*
                                  effectif_df["Nombre total"])/sum(effectif_df["Nombre total"]),2)
    
    effectif_list = effectif_list_h + effectif_list_f 
    age_copep38_adult_list = [x for x in effectif_list if x >= 18]
    nbr_adults =  len(age_copep38_adult_list)
    nbr_enfants = len(effectif_list)-len(age_copep38_adult_list)
    
    age_moyen_copep38_adult = round(mean(age_copep38_adult_list),2)
    _ = plt.hist([effectif_list_h,effectif_list_f], \
                 bins=50, color=['#FFDE59',"#F4511E"] ) #rwidth=0.8,edgecolor='black'
    title = f"Age moyen :{age_moyen_copep38} ans, Age moyen adult :{age_moyen_copep38_adult} ans"
    title = title + "\n"
    title = title + f"Nombre d'adultes {nbr_adults}, Nombre enfants {nbr_enfants}"
    plt.legend(['hommes', 'femmes'])
    plt.ylabel('Nombre de licenciés')
    plt.xlabel('Age')
    #plt.title(title)
    plt.savefig(PATH_OUTPUT / Path('histo.jpeg'))
    dic_stat = {}
    dic_stat['Nbr_clubs'] = len(df_merged)
    nbr_tot = df_merged['Nombre total'].sum()
    dic_stat['Nbr_total'] = nbr_tot
    dic_stat['Nbr_hommes'] = df_merged["Nombre d'hommes"].sum()
    nbr_femmes = df_merged["Nombre de femmes"].sum()
    dic_stat['Nbr_femmes'] = nbr_femmes
    dic_stat['feminisation'] = round(100*nbr_femmes/nbr_tot,2)
    dic_stat['Nbr_jeunes'] = nbr_tot - df_merged["Nombre d'hommes (>=18 ans)"].sum()-df_merged["Nombre dde femmes (>=18 ans)"].sum()
    dic_stat['age_moyen_tot'] = age_moyen_copep38
    dic_stat['age_moyen_adult'] = age_moyen_copep38_adult
    dic_stat['efv'] = len(df_merged.query("efv == 'oui'"))
    dic_stat['Nbr_median_cyclos'] = df_merged['Nombre total'].median()
    
    df_merged = df_merged.fillna(0)
    v = round(df_merged['Age moyen homme'].dot(df_merged["Nombre d'hommes"])/df_merged["Nombre d'hommes"].sum(),2)
    dic_stat['age_moyen_homme'] = v
    v = round(df_merged['Age moyen femme'].dot(df_merged["Nombre de femmes"])/df_merged["Nombre de femmes"].sum(),2)
    dic_stat['age_moyen_femme'] = v
    
    f = []
    efv_d = df_merged.query("efv == 'oui'")
    for idx,row in efv_d.iterrows():
        f.append(dict(id=str(row['ID']).zfill(5),
                      ville=row['Ville du Club'],
                      club=row['Nom du Club']))
    dic_stat['f'] = f
      
    print(PATH_OUTPUT / Path('histo.jpeg'))
    print(root / Path('indicateurs_club.xlsx'))
    
    return df_merged
    
def plot_calendar_genre(df_merged):

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