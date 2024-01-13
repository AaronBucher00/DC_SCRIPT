import tkinter as tk

from import_ifc import import_ifc
from minergie_excel_editor import minergie_excel_editor


#global variabels
glb_exp_data = []


def start_main_func():
    print("INFO - main script started")
    main_window.destroy()


    glb_exp_data = import_ifc() #read Information from IFC


    minergie_excel_editor(glb_exp_data) #edit Information from IFC and write in Minergie-Excel


    print("INFO - main script finished")


if __name__ == '__main__': #main function: starts main function if this script is main
    # Create the main application window
    main_window = tk.Tk()
    main_window.title("Main Window")
    
    # Add a button to open the first window
    open_window1_button = tk.Button(main_window, text="Start Python script", command=start_main_func)
    open_window1_button.pack(pady=10)
    main_window.mainloop()  # Start the main window's main loop