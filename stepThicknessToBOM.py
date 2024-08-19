import PySimpleGUI as sg
from functions import execute_func

# Define a custom theme for the GUI
my_new_theme = {'BACKGROUND': "#31363b",
                'TEXT': "#f9f1ee",
                'INPUT': "#232629",
                'TEXT_INPUT': "#f9f1ee",
                'SCROLL': "#333a41",
                'BUTTON': ('#31363b', '#0dd1fc'),
                'PROGRESS': ('#f9f1ee', '#31363b'),
                'BORDER': 1,
                'SLIDER_DEPTH': 0,
                'PROGRESS_DEPTH': 0}

# Add and set the custom theme
sg.theme_add_new("MyNewTheme", my_new_theme)
sg.theme("MyNewTheme")
sg.set_options(font=("Segoe UI Variable", 11))

logo = r"C:\Users\dominick.cole\Python\stepThicknessToBOM\logos\applogo.ico"
e_logo = r"C:\Users\dominick.cole\Python\stepThicknessToBOM\logos\error.ico"
s_logo = r"C:\Users\dominick.cole\Python\stepThicknessToBOM\logos\success.ico"

# Define the layout for the GUI
layout = [
    [sg.Input(), sg.FolderBrowse("Select Folder to Search", key="source", pad=(5, 5))],
    [sg.Input(), sg.FileBrowse("Select BOM", key="excel", pad=(5, 5))],
    [sg.Button("Get Thicknesses")]
]

# Create the main window with the specified layout
window = sg.Window(title="Step Thickness To BOM", layout=layout, icon=logo)

# Main event loop
while True:
    event, values = window.read()  # Read events and values from the window

    match event:
        case sg.WIN_CLOSED:  # If the window is closed, exit the loop
            break

        case "Get Thicknesses":
            window["Get Thicknesses"].update(disabled=True)

            execute_func(window, values)

        case "Done":
            sg.popup("SUCCESS!\n\n"
                     "Estimated gage thicknesses have been exported into a new BOM"
                     " in the folder you originally selected.\n",
                     custom_text="Exit", icon=s_logo)
            break

        case "Error":
            error_message = values[event]

            if "[Errno 13]" in error_message:
                sg.popup("Ensure that the BOM is closed before running.",
                         custom_text="Exit", icon=e_logo)

            elif "[Errno 2]" in error_message:
                sg.popup("Ensure the correct BOM is selected.",
                         custom_text="Exit", icon=e_logo)

            elif error_message == "no_steps_found":
                sg.popup("No step files found, please select another folder.",
                         custom_text="Exit", icon=e_logo)

            elif error_message == "invalid_BOM":
                sg.popup("Ensure BOM is formatted correctly.\n(Inconsistent column lengths)\n",
                         custom_text="Exit", icon=e_logo)

            else:
                print(values[event])
                sg.popup(f"An unforeseen error occurred.",
                         custom_text="Exit", icon=e_logo)
            break


window.close()  # Close the window
