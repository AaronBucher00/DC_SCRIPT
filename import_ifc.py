import os
import openpyxl
import warnings
import math

import tkinter as tk
from tkinter import filedialog

import ifcopenshell
import ifcopenshell.util
import ifcopenshell.util.element
import ifcopenshell.geom
from ifcopenshell.util.placement import get_local_placement
from ifcopenshell.util.unit import calculate_unit_scale

from export_ifc_data import export_ifc_data

def import_ifc():
    global source_ifc_path  # Declare source_ifc_path as a global variable
    source_ifc_path = ""  # Initialize source_ifc_path as an empty string

    global exp_data  # Declare exp_data as a global variable
    exp_data = []

    root = tk.Tk()
    root.title("Import IFC")

    def load_source_ifc(): #load the source ifc
        global source_ifc_path  # Use the global variable
        filepath = filedialog.askopenfilename(filetypes=[("IFC", "*.ifc")])
        if filepath:
            file_label.config(text=filepath)
            source_ifc_path = filepath 
            save_button.config(state=tk.NORMAL)
            status_label.config(text="IFC ist i.O.", fg="green")
        else:
            status_label.config(text="Fehler: Keine Datei ausgewählt", fg="red")

        warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.worksheet._reader")


    def read_source_ifc():  #get Information from ifc and store in list
        global source_ifc_path  # Use the global variable
        print("INFO - start read IFC...")

        check_file = os.path.isfile(source_ifc_path)

        if(check_file) :
            print("INFO - IFC file is fine")
            
            ifc_file = ifcopenshell.open(source_ifc_path)   
            windows = ifc_file.by_type("IfcWindow")

            ##### SPACE-WINDOWS RELATIONS ##### 
            relations_windows_spaces = []

            for window in windows:
                window_attributes = window.get_info()  # Getting info from attributes

                ifc = ifcopenshell.open(source_ifc_path)  # Get IFC file
                length_unit_factor = calculate_unit_scale(ifc)

                window = ifc.by_guid(window_attributes.get("GlobalId"))  # Get window location
                m4 = ifcopenshell.util.placement.get_local_placement(window.ObjectPlacement)
                location = tuple(map(float, m4[0:3,3] * length_unit_factor))

                tree_settings = ifcopenshell.geom.settings()  # Setup BVH tree
                tree_settings.set(tree_settings.DISABLE_TRIANGULATION, True)
                tree_settings.set(tree_settings.DISABLE_OPENING_SUBTRACTIONS, True)
                it = ifcopenshell.geom.iterator(tree_settings, ifc, include=("IfcSpace",))
                t = ifcopenshell.geom.tree()
                t.add_iterator(it)

                
                try:  # Search tree
                    spaces = t.select(location, extend=0.4)

                    window_info = {  # Create a dictionary for the window
                        "Window_GUID": window_attributes.get("GlobalId"),
                        "Window_Name": window_attributes.get("Name"),
                        "Spaces": []  # List to store space information
                    }
                    
                    for space in spaces:  # Append space information for each space related to the window
                        space_attributes = space.get_info()
                        window_info["Spaces"].append({
                            "Space_GUID": space_attributes.get("GlobalId"),
                            "Space_LongName": space_attributes.get("LongName")
                        })

                    relations_windows_spaces.append(window_info)  # Append the window information to the main list
                except Exception as e:
                    print(f"Error: {e}")


            ##### SPACE PROPERTYS ##### 
            spaces = ifc_file.by_type("IfcSpace")
            m = 0

            for i in spaces:         
                i_attributes = i.get_info() #get all attributes
                exp_data.append([])
                            
                try:
                    exp_data[m].append(i_attributes.get("Name"))
                    exp_data[m].append(i_attributes.get("LongName"))
                except:
                        continue 

                psets = ifcopenshell.util.element.get_psets(i)#getting Info from psets
                exp_data[m].append(psets["BaseQuantities"]["GrossFloorArea"])
                exp_data[m].append(psets["BaseQuantities"]["Height"])

                
                num_windows_in_space = 0  #Number of Windows in the Space
                for window_info in relations_windows_spaces:
                    for space in window_info.get("Spaces", []):
                        space_guid = space.get("Space_GUID")
                        if space_guid == i_attributes.get("GlobalId"):
                            num_windows_in_space += 1
                exp_data[m].append(num_windows_in_space)

                ##### ORIENTATION OF THE WINDOWS ##### 
                def euler_angles_from_matrix_to_three_Letter_z(matrix):  # Get the orientation of the window
                    #Calculate Euler angles from a 3x3 rotation matrix.
                    sy = math.sqrt(matrix[0, 0] * matrix[0, 0] + matrix[1, 0] * matrix[1, 0])
                    singular = sy < 1e-6

                    if not singular:
                        x = math.atan2(matrix[2, 1], matrix[2, 2])
                        y = math.atan2(-matrix[2, 0], sy)
                        z = math.atan2(matrix[1, 0], matrix[0, 0])
                    else:
                        x = math.atan2(-matrix[1, 2], matrix[1, 1])
                        y = math.atan2(-matrix[2, 0], sy)
                        z = 0
                    
                    z_deg = math.degrees(z)

                    if z_deg > -11.25 and z_deg <= 11.25:
                        three_letter_z = "S"
                    elif z_deg > 11.25 and z_deg <= 33.75:
                        three_letter_z = "SSE"
                    elif z_deg > 33.75 and z_deg <= 56.25:
                        three_letter_z = "SE"
                    elif z_deg > 56.25 and z_deg <= 78.75:
                        three_letter_z = "ESE"
                    elif z_deg > 78.75 and z_deg <= 101.25:
                        three_letter_z = "E"
                    elif z_deg > 101.25 and z_deg <= 123.75:
                        three_letter_z = "ENE"
                    elif z_deg > 123.75 and z_deg <= 146.25:
                        three_letter_z = "NE"
                    elif z_deg > 146.25 and z_deg <= 168.75:
                        three_letter_z = "NNE"
                    elif z_deg > 168.75 or z_deg <= -168.75:
                        three_letter_z = "N"
                    elif z_deg <= -146.25 and z_deg > -168.75:
                        three_letter_z = "NNW"
                    elif z_deg <= -123.75 and z_deg > -146.25:
                        three_letter_z = "NW"
                    elif z_deg <= -101.25 and z_deg > -123.75:
                        three_letter_z = "WNW"
                    elif z_deg <= -78.75 and z_deg > -101.25:
                        three_letter_z = "W"
                    elif z_deg <= -56.25 and z_deg > -78.75:
                        three_letter_z = "WSW"
                    elif z_deg <= -33.75 and z_deg > -56.25:
                        three_letter_z = "SW"
                    elif z_deg <= -11.25 and z_deg > -33.75:
                        three_letter_z = "NNW"
                    else:
                        three_letter_z = "H"

                    return three_letter_z

                
                for window_info in relations_windows_spaces:  # Orientation and Dimensions of Windows in the Space
                    for space in window_info.get("Spaces", []):
                        space_guid = space.get("Space_GUID")
                        if space_guid == i_attributes.get("GlobalId"):
                            window_guid = window_info.get("Window_GUID")
                            
                            window_element = ifc.by_guid(window_guid)  # Retrieve the IfcWindow object using its GUID
                            
                            window_psets = ifcopenshell.util.element.get_psets(window_element)  # Getting Info from psets

                            if window_element.ObjectPlacement:  # Append the orientation angles (in degrees) to the list
                                placement_matrix = ifcopenshell.util.placement.get_local_placement(window_element.ObjectPlacement)
                                orientation_matrix = placement_matrix[0:3, 0:3]
                                z_angle = euler_angles_from_matrix_to_three_Letter_z(orientation_matrix)
                                            
                                exp_data[m].append(z_angle)
                            
                            exp_data[m].append(window_psets["BaseQuantities"]["Width"])  # Append window dimensions to the list
                            exp_data[m].append(window_psets["BaseQuantities"]["Height"])
                            exp_data[m].append(window_psets["BaseQuantities"]["Depth"])
                             
                m += 1
        

            export_button = tk.Button(root, text="optional: IFC Daten als Excel exportieren (Speicherort ist Pfad von IFC)",
                                    command=lambda: export_ifc_data(exp_data, source_ifc_path))  # export exp_data as Excel
            export_button.pack(pady=10)

            quit_button = tk.Button(root, text="Weiter", command=root.destroy)  # Quit the Tkinter application
            quit_button.pack(pady=10)

            print("INFO - end of reading IFC")
        
        else:
            status_label.config(text="Fehler: Ungültige Datei", fg="red")


        max_length = max(len(row) for row in exp_data)# Find the length of the longest list in exp_data
        
        for i in range(len(exp_data)):  # Iterate over each row (sublist) and pad with zeros if needed
            exp_data[i] += [0] * (max_length - len(exp_data[i]))

        #print(exp_data) # TEST
        #Raumnummer, Raumname, Fläche, Höhe, Fensteranzahl, Ausrichtung, Fensterbreite, Fensterhöhe, Leibungstiefe

        return (exp_data)
    

    warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.worksheet._reader")

    load_button = tk.Button(root, text="IFC Datei öffnen", command=load_source_ifc)
    load_button.pack(pady=10)

    file_label = tk.Label(root, text="Keine Datei ausgewählt")
    file_label.pack(pady=10)

    save_button = tk.Button(root, text="IFC Datei lesen", command=read_source_ifc, state=tk.DISABLED)
    save_button.pack(pady=10)

    status_label = tk.Label(root, text="")
    status_label.pack(pady=10)

    return (exp_data)