import functions
import PySimpleGUI as sg
import pandas

SUPPORTED_FORMATS = ['xls', 'xlsx']

sg.theme("Black")

file_choice_label = sg.Text("Select desired spreadsheet workbook: ")
file_choice_input = sg.Input(key="file_choice",
                             enable_events=True,
                             visible=True)
file_choice_button = sg.FileBrowse("Choose",
                                   key="filepath",
                                   target="file_choice")

table_choice_label1 = sg.Text("Enter the name of the desired table as it appears")
table_choice_label2 = sg.Text("on the tab in the workbook:")
table_choice_input = sg.InputText(tooltip="table name", key="table")

output_dir_label = sg.Text("Pick the desired destination directory:")
output_dir_input = sg.Input(key="folder_choice",
                            enable_events=True,
                            visible=True)
output_dir_button = sg.FolderBrowse("Choose",
                                    key="dest_folder",
                                    target="folder_choice")

output_file_label = sg.Text("Enter the name of the text file output:")
output_file_choice = sg.InputText(default_text="output",
                                  key="filename")

success_label = sg.Text(key="completed", text_color="white")

convert_button = sg.Button("Convert Table")
quit_button = sg.Button("Quit", key="quit")

layout_col1 = [[file_choice_label],
               [file_choice_input],
               [file_choice_button],
               [table_choice_label1],
               [table_choice_label2],
               [table_choice_input]]

layout_col2 = [[output_dir_label],
               [output_dir_input],
               [output_dir_button],
               [output_file_label],
               [output_file_choice],
               [convert_button, quit_button]]

col1 = sg.Column(layout=layout_col1)
col2 = sg.Column(layout=layout_col2)

window = sg.Window("Table Converter",
                   layout=[[col1, col2],
                           [success_label]])

while True:
    event, values = window.read()
    print(event, values)
    window["folder_choice"].update(value=values["dest_folder"])
    window["file_choice"].update(value=values["filepath"])
    match event:
        case sg.WIN_CLOSED:
            break
        case "quit":
            break


window.close()