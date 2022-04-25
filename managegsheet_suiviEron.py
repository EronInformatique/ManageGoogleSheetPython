import pygsheets
def create_duplicate_sheets(gsheet,dic_sheet_to_create,wksheet_to_duplicate):
    "duplicate or create worksheet"

    months=list(dic_sheet_to_create.keys())
    categories_formation = ["INF","KIN","MED","DEN","PHA","INT"]
    depart_sess= list(dic_sheet_to_create["02"].keys())

    new_worksheet_created = {month: {dep:{cat:[] for cat in categories_formation } for dep in depart_sess} for month in months}
    for month in dic_sheet_to_create.keys():
        for dep in  depart_sess:
            for cat in categories_formation:
                if dic_sheet_to_create[month][dep][cat]==[]:
                    continue
                else:
                    try:
                        print(dic_sheet_to_create[month][dep][cat][0])
                        gsheet.worksheet_by_title(dic_sheet_to_create[month][dep][cat][0])
                        new_worksheet_created[month][dep][cat].append(gsheet.worksheet_by_title(dic_sheet_to_create[month][dep][cat][0]))
                        # new_worksheet_created[month][dep][cat][0].update_value("M1",f'="{dep}/{month}"')
                        # new_worksheet_created[month][dep][cat][0].update_value("A2","(Session du "+dep+"/"+month+"/2022 -  ) "+cat)
                        # new_worksheet_created[month][dep][cat][0].clear(start="J7",end="L1000")
                        # new_worksheet_created[month][dep][cat][0].clear(start="Q7",end="Q1000")
                    except pygsheets.WorksheetNotFound as error:
                        print(error)
                        print(dic_sheet_to_create[month][dep][cat][0])
                        gsheet.add_worksheet(dic_sheet_to_create[month][dep][cat][0],src_worksheet =wksheet_to_duplicate['02']['01'][cat][0])
                        new_worksheet_created[month][dep][cat].append(gsheet.worksheet_by_title(dic_sheet_to_create[month][dep][cat][0]))
                        new_worksheet_created[month][dep][cat][0].update_value("M1",f'="{dep}/{month}"')
                        new_worksheet_created[month][dep][cat][0].update_value("A2","(Session du "+dep+"/"+month+"/2022 -  )"+cat)
                        new_worksheet_created[month][dep][cat][0].clear(start="J7",end="L1000")
                        new_worksheet_created[month][dep][cat][0].clear(start="Q7",end="Q1000")
                        new_worksheet_created[month][dep][cat][0].clear(start="E7",end="E1000")
                        new_worksheet_created[month][dep][cat][0].clear(start="H7",end="H1000")
                        new_worksheet_created[month][dep][cat][0].clear(start="I7",end="I1000")
                        pass

    for month in new_worksheet_created.keys():
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
    n=1
    nb_max=len(new_worksheet_created.keys())*len(depart_sess)*len(categories_formation)
    skip=True
    for month in new_worksheet_created.keys():
        for dep in  depart_sess:
            for cat in categories_formation:
                if new_worksheet_created[month][dep][cat]!= []:
                    ele = new_worksheet_created[month][dep][cat][0]
                    range_name="Depart"+cat+dep+month
                    range_name_start="DepartStatutCRM"+cat+dep+month
                    # if ele.get_named_ranges() != []:
                    named_ranges = ele.get_named_ranges()
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
                    continue
                    

def update_accueil_sheet(gsheet,dic_sheet_to_create):
    """mise à jour de l'acceuil"""

    months=list(dic_sheet_to_create.keys())
    categories_formation = ["INF","KIN","MED","DEN","PHA","INT"]
    depart_sess= list(dic_sheet_to_create["02"].keys())

    # for month in months:
    #     for dep in depart_sess:
    dic_session_start_colH={dep+month:[] for month in months for dep in depart_sess if dep+month!="0102"}
    dic_session_start_colI={dep+month:[] for month in months for dep in depart_sess if dep+month!="0102"}

    list_depart=[]
    list_formula_h=[]
    list_formula_i=[]

    for month in months :
        for dep in depart_sess:
            if dep+month!="0102":
                list_depart.append("Départ "+dep+"/"+month+"/"+"2022")

    for monthdep in dic_session_start_colH.keys():
        for cat in categories_formation:
            dic_session_start_colH[monthdep].append(f'ARRAYFORMULA(SOMME(--REGEXMATCH(Depart{cat}{monthdep};"^*{cat}*")))-ARRAYFORMULA(SOMME(--REGEXMATCH(DepartStatutCRM{cat}{monthdep};"Excusé")))')
            dic_session_start_colI[monthdep].append(f'ARRAYFORMULA(SOMME(--REGEXMATCH(ARRAYFORMULA(Depart{cat}{monthdep}&" "&DepartStatutCRM{cat}{monthdep});ARRAYFORMULA(CONCATENER("^*";"{cat}";".*";"Present";'+r'"\b"'+")))))")
        list_formula_h.append("="+"+".join(dic_session_start_colH[monthdep]))
        list_formula_i.append("="+"+".join(dic_session_start_colI[monthdep]))
                # for row in range(42,80,2):
                #     print(row)
    wk=gsheet.worksheet_by_title("(00) Accueil")
    for row,formula_h,formula_i,depart in zip(range(42,46,2),list_formula_h[:2],list_formula_i[0:2],list_depart[0:2]):
        wk.update_value((row,6),depart)
        wk.update_value((row,8),formula_h)
        wk.update_value((row,9),formula_i)
        wk.update_value((row,3),depart)
        
        print("Ajout de formule dans page Acceuil: "+ formula_h + " "+formula_i+ " "+depart)
        print(row)


    stop_here=0



if __name__ == "__main__":
    gc = pygsheets.authorize(client_secret='/Users/acapai/Documents/Git/Espace-test/Scrapping-With-Python/ManageGoogleSheetPython/code_secret_client_173621592213-0llal3ntmv316usvtboglpb52leq6jcu.apps.googleusercontent.com.json')
    sh = gc.open_by_key('13VqSH8KjAzB3-mroVhtUJjXgO2Gs31UtpqdehiLMyRs')
    # gc.set_batch_mode(False)
    months=["02","03","04","05","06","07","08","09","10","11","12"]
    year = ["2022"]
    categories_formation = ["INF","KIN","MED","DEN","PHA","INT"]
    depart_sess=["01","15"]

    dic_sheets_to_create= {month: {dep:{cat:[] for cat in categories_formation } for dep in depart_sess} for month in months}
    for month in months :
        for dep in depart_sess:
            for cat in categories_formation:
                if month == "02" and dep=="15":
                    dic_sheets_to_create[month][dep][cat].append("(Départ "+dep+"/"+month+"/"+year[0]+") "+cat)
                    print("(Départ "+dep+"/"+month+"/"+year[0]+") "+cat)
                elif month != "02":
                    dic_sheets_to_create[month][dep][cat].append("(Départ "+dep+"/"+month+"/"+year[0]+") "+cat)
                    print("(Départ "+dep+"/"+month+"/"+year[0]+") "+cat)

    # wksheet_to_duplicate = []
    el = "02"
    wksheet_to_duplicate={el:{dep:{cat:[] for cat in categories_formation } for dep in depart_sess}}
    for cat in categories_formation:
        print("(Départ "+dep+"/"+el+"/"+year[0]+") "+cat)
        wksheet_to_duplicate[el]["01"][cat].append(sh.worksheet('title',"(Départ "+"01"+"/"+el+"/"+year[0]+") "+cat))
# sheet_to_delete = []
    # for wk in range(22,len(sh._sheet_list)):
    #     sheet_to_delete.append(sh[wk])
    # for sheet in sheet_to_delete:
    #     sh.del_worksheet(sheet)
    stop= True
    # update_accueil_sheet(sh,dic_sheets_to_create)

    
    create_duplicate_sheets(sh,dic_sheets_to_create,wksheet_to_duplicate)