import os
import pandas as pd

def export_ifc_data(exp_raw_data, source_ifc_path):

    if len(exp_raw_data) > 0:

        status_label = None

        def ifc_path_to_export_folder(input_string):
            index = input_string.rfind("/")
            if index != -1:
                result = input_string[:index]  # Abschneiden bis zum gefundenen Index (einschließlich)
                return result
            else:
                return input_string  # Wenn das Trennzeichen nicht gefunden wird, gib den ursprünglichen String zurück

        export_folder = ifc_path_to_export_folder(source_ifc_path)
        export_folder_a = "C:/Users/abu/Downloads/HSLU/PYTHON_scripts/DC_SCRIPT/Semesterprojekt/240104"
        xlsx_file_name = "raw_ifc_data_Minergie-Excel.xlsx"
        exp_path = os.path.join(export_folder, xlsx_file_name)

        check_file = os.path.isfile(exp_path)

        if(not(check_file)):
            pd.DataFrame([]).to_excel(exp_path)

        col_label = [ #cloumn titles
            "Raumnummer", "Raumname", "Raumfläche", "Raumhöhe", "Fensteranzahl",
            "Ausrichtung FE01", "Fensterbreite FE01", "Fensterhöhe FE01", "Leibungstiefe FE01",
            "Ausrichtung FE02", "Fensterbreite FE02", "Fensterhöhe FE02", "Leibungstiefe FE02",
            "Ausrichtung FE03", "Fensterbreite FE03", "Fensterhöhe FE03", "Leibungstiefe FE03",
            "Ausrichtung FE04", "Fensterbreite FE04", "Fensterhöhe FE04", "Leibungstiefe FE04" ]
            
        df = pd.DataFrame(exp_raw_data, columns=col_label)
                
        with pd.ExcelWriter(exp_path, engine='openpyxl', if_sheet_exists='replace', mode='a') as writer:  
            df.to_excel(writer, sheet_name="exp_raw_data", index=False)
           
           
        if status_label:
            status_label.config(text="Excel abgespeichert\n" + export_folder, fg="green")


        print("INFO - Excel exported")