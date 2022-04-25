import pygsheets
import os
import pickle
import io


def load_wk_sheet_structure_suivi(new_start):
    """Creation d'un dossier data + sauvegarde d'un fichier pickle pour récupérer plus rapidement la structure des sheets"""
    
    try:
        path_directory=os.path.join(os.path.dirname(__file__),"data")
        print(path_directory)
        os.makedirs(path_directory)
    except FileExistsError:
    # directory already exists
        pass 
    
    file_pickle_path=os.path.join(path_directory,"data_wk_sheet.pickle")
    if os.path.exists(file_pickle_path) and new_start==True:
        os.remove(file_pickle_path)
    else:
        print("Ok") 

    file_to_write = open(file_pickle_path, "wb")
    try:
        dic_sheet_structure_suivi=pickle.load(file_to_write)
        return dic_sheet_structure_suivi
    except io.UnsupportedOperation as e:
        print(e)
        months=["02","03","04","05","06","07","08","09","10","11","12"]
        year = ["2022"]
        categories_formation = ["INF","KIN","MED","DEN","PHA","INT"]
        depart_sess=["01","15"]

        dic_sheet_structure_suivi= {month: {dep:{cat:[] for cat in categories_formation } for dep in depart_sess} for month in months}
        for month in months :
            for dep in depart_sess:
                for cat in categories_formation:
                    # if month == "02" and dep=="15":
                    dic_sheet_structure_suivi[month][dep][cat].append("(Départ "+dep+"/"+month+"/"+year[0]+") "+cat)
                    # print("(Départ "+dep+"/"+month+"/"+year[0]+") "+cat)
                    # elif month != "02":
                    #     dic_sheet_structure_suivi[month][dep][cat].append("(Départ "+dep+"/"+month+"/"+year[0]+") "+cat)
                    #     print("(Départ "+dep+"/"+month+"/"+year[0]+") "+cat)
        
        pickle.dump(dic_sheet_structure_suivi, file_to_write)
        return dic_sheet_structure_suivi

def duplicate_wk_sheet_from_exist(template_use, start_wk_sheet,dic_sheet_structure,All):
    """Dupliquer le ou les sheet souhaité"""

    year=["2022"]
    categories_formation = ["INF","KIN","MED","DEN","PHA","INT"]
    # categories_formation = ["INF"]   
    
    
    jour_start = str(start_wk_sheet[0])
    mois_start = str(start_wk_sheet[1])

    jour=  str(template_use[0])
    mois = str(template_use[1])

    if All:
        months=list(dic_sheet_structure.keys())
        depart_sess= list(dic_sheet_structure[mois_start].keys())
        depart_sess_duplicate=[list(dic_sheet_structure[mois_start].keys())]
    else:
        months=[mois_start]
        depart_sess=[jour_start]
        depart_sess_duplicate=[jour]



    wksheet_to_duplicate={mois:{dep:{cat:[] for cat in categories_formation } for dep in depart_sess_duplicate}}
    for cat in categories_formation:
        print("(Départ "+jour+"/"+mois+"/"+year[0]+") "+cat)
        wksheet_to_duplicate[mois][jour][cat].append(sh.worksheet('title',"(Départ "+jour+"/"+mois+"/"+year[0]+") "+cat))
    
    new_worksheet_created = {month: {dep:{cat:[] for cat in categories_formation } for dep in depart_sess} for month in months}
    for month in months:
        for dep in  depart_sess:
            for cat in categories_formation:
                # if dic_sheet_structure[month][dep][cat]==[]:
                #     continue
                # else:
                try:
                    print(dic_sheet_structure[month][dep][cat][0])
                    sh.worksheet_by_title(dic_sheet_structure[month][dep][cat][0])
                    new_worksheet_created[month][dep][cat].append(sh.worksheet_by_title(dic_sheet_structure[month][dep][cat][0]))
                except pygsheets.WorksheetNotFound as error:
                    print(error)
                    print(dic_sheet_structure[month][dep][cat][0])
                    sh.add_worksheet(dic_sheet_structure[month][dep][cat][0],src_worksheet =wksheet_to_duplicate[mois][jour][cat][0])
                    new_worksheet_created[month][dep][cat].append(sh.worksheet_by_title(dic_sheet_structure[month][dep][cat][0]))
                    new_worksheet_created[month][dep][cat][0].update_value("M1",f'="{dep}/{month}"')
                    new_worksheet_created[month][dep][cat][0].update_value("A2","(Session du "+dep+"/"+month+"/"+year[0]+" -  )"+cat)
                    new_worksheet_created[month][dep][cat][0].clear(start="J7",end="L1000")
                    new_worksheet_created[month][dep][cat][0].clear(start="Q7",end="Q1000")
                    # new_worksheet_created[month][dep][cat][0].clear(start="E7",end="E1000")
                    # new_worksheet_created[month][dep][cat][0].clear(start="H7",end="H1000")
                    new_worksheet_created[month][dep][cat][0].clear(start="I7",end="I1000")
                    pass

    return new_worksheet_created

def remove_named_range(new_worksheet_created,start_wk_sheet):
    """Remove range named"""
    
    # jour_start = str(start_wk_sheet[0])
    # year=["2022"]
    mois_start = str(start_wk_sheet[1])
    categories_formation = ["INF","KIN","MED","DEN","PHA","INT"]    
    # categories_formation = ["INF"]   
    months=list(new_worksheet_created.keys())
    depart_sess= list(new_worksheet_created[mois_start].keys())


    for month in months:
        for dep in  depart_sess:
            for cat in categories_formation:
                if new_worksheet_created[month][dep][cat]!= []:
                    ele = new_worksheet_created[month][dep][cat][0]
                    if ele.get_named_ranges() != []:
                        named_ranges = ele.get_named_ranges()
                        range_name="Depart"+cat+dep+month
                        range_name_start="DepartStatutCRM"+cat+dep+month
                        for name_range in named_ranges:
                            if name_range.name == range_name or name_range.name == range_name_start:
                                continue
                            else:
                                ele.delete_named_range(name_range.name)
                                print("Delete Range name: "+name_range.name)
                    else:
                        continue
                else:
                    continue

def set_named_range(new_worksheet_created,start_wk_sheet):
    """Nommer les plages a utiliser dans la page accueil"""

    mois_start = str(start_wk_sheet[1])
    categories_formation = ["INF","KIN","MED","DEN","PHA","INT"]
    # categories_formation = ["INF"]   
    months=list(new_worksheet_created.keys())
    depart_sess= list(new_worksheet_created[mois_start].keys())

    n=1
    nb_max=len(new_worksheet_created.keys())*len(depart_sess)*len(categories_formation)
    skip=True
    for month in months:
        for dep in  depart_sess:
            for cat in categories_formation:
                if new_worksheet_created[month][dep][cat]!= []:
                    ele = new_worksheet_created[month][dep][cat][0]
                    range_name="Depart"+cat+dep+month
                    range_name_start="DepartStatutCRM"+cat+dep+month
                    # if ele.get_named_ranges() != []:
                    named_ranges = ele.get_named_ranges()
                    if named_ranges != []:
                        for name_range in named_ranges:
                            if name_range.name == range_name or name_range.name == range_name_start:
                                continue
                            else:
                                grange =  pygsheets.GridRange(worksheet=ele,start=(7,13),end=(1000,13))
                                ele.create_named_range(range_name,(7,13),(1000,13),grange)
                                print("Created Range name: "+range_name+f" ({n}/{nb_max})")
                                grange_statut =  pygsheets.GridRange(worksheet=ele,start=(7,15),end=(1000,15))
                                ele.create_named_range(range_name_start,(7,15),(1000,15),grange_statut)
                                print("Created Range name: "+range_name_start+f" ({n}/{nb_max})")
                                n=n+1
                        # else:
                        #     continue
                    else:
                        grange =  pygsheets.GridRange(worksheet=ele,start=(7,13),end=(1000,13))
                        ele.create_named_range(range_name,(7,13),(1000,13),grange)
                        print("Created Range name: "+range_name+f" ({n}/{nb_max})")
                        grange_statut =  pygsheets.GridRange(worksheet=ele,start=(7,15),end=(1000,15))
                        ele.create_named_range(range_name_start,(7,15),(1000,15),grange_statut)
                        print("Created Range name: "+range_name_start+f" ({n}/{nb_max})")
                        n=n+1

                else:
                    continue


def update_wk_sheet_accueil(new_worksheet_created,start_wk_sheet,dep_row_num,stats_formations_row,stats_formations_col):
    """Mise à jour de l'acceuil"""
    year=["2022"]
    mois_start = str(start_wk_sheet[1])
    categories_formation = ["INF","KIN","MED","DEN","PHA","INT"]     
    months=list(new_worksheet_created.keys())
    depart_sess= list(new_worksheet_created[mois_start].keys())

    col_numbers = [ord(letter[0]) - 96 for letter in stats_formations_col]

    wk=sh.worksheet_by_title("(00) Accueil")
    # for month in months:
    #     for dep in depart_sess:
    dic_session_start_colH={dep+month:[] for month in months for dep in depart_sess} # if dep+month!="0102"
    dic_session_start_colI={dep+month:[] for month in months for dep in depart_sess}


    dic_stat_cat_colS={cat:[] for cat in categories_formation} # if dep+month!="0102"
    dic_stat_cat_colT={cat:[] for cat in categories_formation}

    to_add_formula_s={cat:[] for cat in categories_formation}
    to_add_formula_t={cat:[] for cat in categories_formation}
    # wk.unlink()
    for idx,cat in enumerate(categories_formation):
        cell_R=stats_formations_col[0].upper()+str(stats_formations_row[idx])
        dic_stat_cat_colS[cat].append(wk.cell(cell_R).formula)
        cell_S=stats_formations_col[1].upper()+str(stats_formations_row[idx])
        dic_stat_cat_colT[cat].append(wk.cell(cell_S).formula)
        print("Récupération des formules colonne cellules: "+cell_R + " et " + cell_S )
    # wk.link()
    # for month in months:
    #     for dep in  depart_sess:
    #         for cat in categories_formation:
    #             if new_worksheet_created[month][dep][cat]!= []:
    for idx,cat in enumerate(categories_formation):
        for month in months:
            for dep in  depart_sess:
                range_name="Depart"+cat+dep+month
                range_name_start="DepartStatutCRM"+cat+dep+month
                print("Ajout des formules dans le dictionnaire: catégorie: " + cat + " ,ligne numéro: " + str(stats_formations_row[idx]))
                to_add_formula_s[cat].append(f'ARRAYFORMULA(SOMME(--REGEXMATCH({range_name};ARRAYFORMULA(CONCATENER(".*?\\b";SUBSTITUE($K{stats_formations_row[idx]};" ";"");"\\b")))))-ARRAYFORMULA(SOMME(--REGEXMATCH(ARRAYFORMULA({range_name}&" "&{range_name_start});ARRAYFORMULA(CONCATENER(".*?\\b";SUBSTITUE($K{stats_formations_row[idx]};" ";"");"\\b.*?\\b";"Excus";"\\b")))))')
                to_add_formula_t[cat].append(f'ARRAYFORMULA(SOMME(--REGEXMATCH(ARRAYFORMULA({range_name}&" "&{range_name_start});ARRAYFORMULA(CONCATENER(".*?\\b";SUBSTITUE($K{stats_formations_row[idx]};" ";"");"\\b.*?\\b";"Present";"\\b")))))')

    wk.unlink()
    for idx,cat in enumerate(categories_formation):
        cell_R=stats_formations_col[0].upper()+str(stats_formations_row[idx])
        cell_S=stats_formations_col[1].upper()+str(stats_formations_row[idx])

        new_formula_s = dic_stat_cat_colS[cat][0]+"+"+"+".join(to_add_formula_s[cat])
        new_formula_t = dic_stat_cat_colT[cat][0]+"+"+"+".join(to_add_formula_t[cat])
        wk.update_value(cell_R,new_formula_s)
        wk.update_value(cell_S,new_formula_t)
        print("Mise à jour des cellules: "+cell_R + " et " + cell_S )

    wk.link()


    list_depart=[]
    list_formula_g=[]
    list_formula_h=[]

    for month in months :
        for dep in depart_sess:
            # if dep+month!="0102":
            list_depart.append("Départ "+dep+"/"+month+"/"+year[0])

    for monthdep in dic_session_start_colH.keys():
        for cat in categories_formation:
            if cat == "INT":
                cat_bis="PSY"
            else:
                cat_bis = cat
            dic_session_start_colH[monthdep].append(f'ARRAYFORMULA(SOMME(--REGEXMATCH(Depart{cat}{monthdep};"^*{cat_bis}*")))-ARRAYFORMULA(SOMME(--REGEXMATCH(DepartStatutCRM{cat}{monthdep};"Excus")))')
            dic_session_start_colI[monthdep].append(f'ARRAYFORMULA(SOMME(--REGEXMATCH(ARRAYFORMULA(Depart{cat}{monthdep}&" "&DepartStatutCRM{cat}{monthdep});ARRAYFORMULA(CONCATENER("^*";"{cat_bis}";".*";"Present";'+r'"\b"'+")))))")
        list_formula_g.append("="+"+".join(dic_session_start_colH[monthdep]))
        list_formula_h.append("="+"+".join(dic_session_start_colI[monthdep]))
                # for row in range(42,80,2):
                #     print(row)

    end_row_num = dep_row_num+len(list_depart)+3

    wk.unlink()
    for row,formula_h,formula_i,depart in zip(range(dep_row_num,end_row_num,2),list_formula_g,list_formula_h,list_depart):
        # wk.update_value((row,6),depart)
        wk.update_value((row,7),formula_h)
        wk.update_value((row,8),formula_i)
        wk.update_value((row,3),depart)
        print("Ajout de formule dans page Acceuil: "+depart)
        print(row) 
    wk.link()     






if __name__ == "__main__":    
    new_start=False
    dic_sheet_structure_suivi= load_wk_sheet_structure_suivi(new_start)
    # Replace by absolute Path gc = pygsheets.authorize(client_secret='/Users/acapai/Desktop/Dossiers_Important/Secure/informatique mail/code_secret_client_173621592213-0llal3ntmv316usvtboglpb52leq6jcu.apps.googleusercontent.com.json')
    gc = pygsheets.authorize(client_secret='./code_secret_client_173621592213-0llal3ntmv316usvtboglpb52leq6jcu.apps.googleusercontent.com.json')
    sh = gc.open_by_key('13VqSH8KjAzB3-mroVhtUJjXgO2Gs31UtpqdehiLMyRs')

    template_use = ["01","02"] # 1er element Jour , second Mois (quel feuille du google sheet sert de template)
    start_wk_sheet = ["01","04"]  # 1er element Jour , second Mois (à partir de quelle feuille on duplique)
    
    # 1) Dupliquer les sheets pour les nouveaux départ de session
    new_worksheet_created=duplicate_wk_sheet_from_exist(template_use, start_wk_sheet,dic_sheet_structure_suivi,All=False)

    # 2) Supprimer les noms de plages protégées qui ont été duppliquées
    remove_named_range(new_worksheet_created,start_wk_sheet)

    # 3) Nommer les noms de plages protégées qui ont été duppliquées
    set_named_range(new_worksheet_created,start_wk_sheet)

    # 4) Update Page accueil
    dep_row_num=48
    stats_formations_row= [7,37,67,95,111,121]
    stats_formations_col= ["r","s"]
    update_wk_sheet_accueil(new_worksheet_created,start_wk_sheet,dep_row_num,stats_formations_row,stats_formations_col)



