import pygsheets
from fuzzywuzzy import fuzz
def create_duplicate_sheets(gsheet,dic_name_shift):
    "duplicate or create worksheet"
    # names_ranges = gsheet.named_ranges
    # id_spreadsheet = gsheet.id

    wksheet_to_duplicate = gsheet.worksheet('title',"NEW - Rapport commerciaux")
    # id_wksheet = wksheet_to_duplicate.id

    new_worksheet_created = []
    for new_wksheet in dic_name_shift.keys():
        try:
            new_worksheet_created.append(gsheet.worksheet('title',"Rapport commerciaux-"+new_wksheet))
        except pygsheets.WorksheetNotFound as error:
            print(error)
            gsheet.add_worksheet("Rapport commerciaux-"+new_wksheet)#,src_worksheet =wksheet_to_duplicate)
            new_worksheet_created.append(gsheet.worksheet('title',"Rapport commerciaux-"+new_wksheet))
            # for wk in new_worksheet_created:
            #     named_ranges = wk.get_named_ranges()
            #     for name_range in named_ranges:
            #         wk.delete_named_range(name_range.name)
            pass

    # for ele in new_worksheet_created:
    #     if ele.get_named_ranges() != []:
    #         named_ranges = ele.get_named_ranges()
    #         for name_range in named_ranges:
    #             ele.delete_named_range(name_range.name)

    for (key_dic,item),(idx_list,ele) in zip(dic_name_shift.items(),enumerate(new_worksheet_created)):
        nb_col = ele.cols
        if ele.get_named_ranges() == []:
            for col in range(nb_col):
                grange =  pygsheets.GridRange(worksheet=ele,start=(1,col+1),end=(1000,col+1))
                if col == 0:
                    ele.create_named_range("Commerciaux"+key_dic,(1,1),(1000,1),grange)
                if col !=0:
                    ele.create_named_range(item[col-1],(1,col+1),(1000,col+1),grange)
        else:
            continue
    # dic_list_named_ranges= {wk.title:wk.get_named_ranges() for wk in new_worksheet_created}

    
    wksheet_to_duplicate_2 = gsheet.worksheet('title',"NEW - Tableau Team Eron 2022")
    new_worksheet_tableau_team_eron = []
    for new_wksheet in dic_name_shift.keys():
        try:
            new_worksheet_tableau_team_eron.append(gsheet.worksheet('title',"NEW - Tableau Team Eron "+new_wksheet))
        except pygsheets.WorksheetNotFound as error:
            print(error)
            gsheet.add_worksheet("NEW - Tableau Team Eron "+new_wksheet,src_worksheet =wksheet_to_duplicate_2)
            new_worksheet_tableau_team_eron.append(gsheet.worksheet('title',"NEW - Tableau Team Eron "+new_wksheet))
            pass
    
    start_col_present = 6
    start_col_prospect = 17
    start_col_anexre = 28
    for wk,(key_dic,list_named_range) in zip(new_worksheet_tableau_team_eron,dic_name_shift.items()):
        commerciaux_range = "Commerciaux"+key_dic
        for idx,name in enumerate(list_named_range):
            if "PresInf" in name:
                start_col = start_col_present
            elif "ProspectInf" in name:
                start_col = start_col_prospect
            elif "AnExReInf" in name:
                start_col = start_col_anexre
            wk.update_value((9,start_col), f"=ArrayFormula(SOMME.SI.ENS({name};regexmatch({commerciaux_range};$E9);VRAI))")
            start_col=start_col+1



if __name__ == "__main__":
    gc = pygsheets.authorize(client_secret='/Users/acapai/Documents/Git/Espace-test/Scrapping-With-Python/ManageGoogleSheetPython/code_secret_client_173621592213-0llal3ntmv316usvtboglpb52leq6jcu.apps.googleusercontent.com.json')
    sh = gc.open_by_key('1L1h9qKd_fz4xiVjN_FFWxvByWHRJwmmB6LIi6vQ-A5A')
    months=["Jan","Fev","Mars","Avr","Mai","Juin","Jui","Aout","Sep","Oct","Nov","Dec"]
    year = ["2022"]
    categories_formation = ["Inf","Kin","Med","Den","Pha"]
    type_tag=["","New"]
    type_status = ["Pres","Prospect","AnExRe"]

    start_cell = ["A","B","C","D","E","F","G","H","I"] 
    dic_sheets_by_months= {month+year[0]:[] for month in months}
    for keydic in dic_sheets_by_months.keys():
        for type in type_status:
            for cat in categories_formation:
                for tag in type_tag:
                    dic_sheets_by_months[keydic].append(type+tag+cat+keydic)

    
    create_duplicate_sheets(sh,dic_sheets_by_months)
