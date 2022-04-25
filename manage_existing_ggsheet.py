import pygsheets

def clear_column_by_title(gc,wk_sheet_list,column_list):
    """Clear column by shee title"""
    
    for wk in wk_sheet_list:
        gc.set_batch_mode(True)
        for col in column_list:
            wk.clear(start=col[0],end=col[1])
            print("Delete col: "+ ":".join(col[:]) + " from sheet: "+wk.title )
        gc.run_batch() # All the above requests are executed here
        gc.set_batch_mode(False)


def copy_paste_values_from_range_to_another(wk_sheet_list,column_list_infos_to_copy,column_letter_to_paste):
    """Copy column with non empty value"""

    for wk in wk_sheet_list:
        if len(column_list_infos_to_copy) == len(column_letter_to_paste):
            gc.set_batch_mode(True)
            for col,col_to_paste in zip(column_list_infos_to_copy,column_letter_to_paste):
                letter_col=col[0]
                start_row=col[1]
                end_row=col[2]

                letter_col_to_paste=col_to_paste[0]

                matrix=wk.get_values(letter_col+start_row,letter_col+end_row)
                print("Col Copied from: "+ letter_col+start_row+ ":" +letter_col+end_row + " in col range: "+letter_col_to_paste+start_row+":"+letter_col_to_paste+str(len(matrix)+int(start_row)-1) + " from sheet" + wk.title )
                wk.update_values(crange=letter_col_to_paste+start_row+":"+letter_col_to_paste+str(len(matrix)+int(start_row)-1),values=matrix)
            gc.run_batch() # All the above requests are executed here
            gc.set_batch_mode(False)

if __name__ == "__main__":
    gc = pygsheets.authorize(client_secret='/Users/acapai/Documents/Git/Espace-test/Scrapping-With-Python/ManageGoogleSheetPython/code_secret_client_173621592213-0llal3ntmv316usvtboglpb52leq6jcu.apps.googleusercontent.com.json')
    sh = gc.open_by_key('1YWIUyO52XMBa1c6RpJmY7elMEbNWfP-LCYlZqd5aElE')

    ## Choix de la liste de feuilles concerné
    wk_sheet_list_clear_static_column=[]
    keep_wk = False
    for wk in sh._sheet_list:
        if keep_wk is True:
            wk_sheet_list_clear_static_column.append(wk)
            continue
        
        if wk.title == "(Départ 15/03/2022) KIN" :
            wk_sheet_list_clear_static_column.append(wk)
            keep_wk = True
        else:
            continue
    column_list =[("E7","E1000"),("H7","H1000"),("I7","I1000")]
    clear_column_by_title(gc,wk_sheet_list_clear_static_column,column_list)

    column_list_infos_to_copy =[("D","7","1000"),("G","7","1000")]
    column_letter_to_paste =["E","H"]
    copy_paste_values_from_range_to_another(wk_sheet_list_clear_static_column,column_list_infos_to_copy,column_letter_to_paste)

