import functions
import PySimpleGUI as sg
import pandas

SUPPORTED_FORMATS = ['xls', 'xlsx']

sg.theme("DarkPurple")

file_choice_label = sg.Text("Select desired spreadsheet workbook: ")
file_choice_input = sg.Input()
file_choice_button = sg.FileBrowse("Choose", key="filepath")

table_choice_label = sg.Text("Enter the name of the desired table as it "
                              "appears on the tab in the workbook:")
table_choice_input = sg.InputText(tooltip="table name", key="table")

output_dir_label = sg.Text("Pick the desired destination directory:")
output_dir_input = sg.Input()
output_dir_button = sg.FolderBrowse("Choose", key="dest_folder")

output_file_label = sg.Text("Enter the name of the text file output:")
output_file_choice = sg.InputText(default_text="output",
                                  key="filename")

convert_button = sg.Button("Convert Table")

layout_col1 = [[file_choice_label],
               [file_choice_input],
               [file_choice_button]]

layout_col2 = [[output_file_label],
               [output_dir_input],
               [output_dir_button]]

col1 = sg.Column(layout=layout_col1)
col2 = sg.Column(layout=layout_col2)

window = sg.Window("Table Converter",
                   layout=[[col1, col2],
                           [table_choice_label],
                           [table_choice_input, convert_button]])

window.read()
window.close()