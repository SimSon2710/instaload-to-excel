import PySimpleGUI as sg
import sys
from os.path import join, exists, dirname
from datetime import datetime


class GUI:
    sg.theme('Tan Blue')
    path_len = len(__file__)
    folder_input_len = (int(path_len * 1), 1)
    file_input_len = (int(path_len * 1.35), 1)
    total_jpg = 0
    jpg_counter = 0

    def __init__(self):
        self.values = None
        self.event = None
        self.window = None

    def browse_folder_layout(self):
        default_path = join(__file__, "loads") if exists(path=join(__file__, "loads")) else dirname(dirname(__file__))
        layout = [[sg.Text('Select the folder containing the Instaloader Files (*.json.xz).')],
                  [sg.Text('Folder:', size=(5, 1)), sg.Input(default_text=default_path, size=self.folder_input_len),
                   sg.FolderBrowse(button_text='Search', initial_folder=default_path)],
                  [sg.Submit(button_text="Ok"), sg.Cancel(button_text="Cancel")],
                  [sg.Text("If there are several folders, move them to a single folder and select it.",
                           text_color="black")]
                  ]

        self.window = sg.Window('Select Instaloader-Folder', layout)

    def save_to_file_layout(self):
        cur_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        default_path = join(dirname(__file__), f"Instaload_{cur_time}.xlsx")
        layout = [[sg.Text('Where do you want to save the EXCEL-File?')],
                  [sg.Text('Folder:', size=(9, 1)), sg.Input(default_text=default_path,
                   size=self.file_input_len), sg.FileSaveAs(button_text='Speichern unter',
                   initial_folder=default_path, default_extension='xlsx', file_types=(('EXCEL', 'XLSX'),))],
                  [sg.Submit(button_text="Finish"), sg.Cancel(button_text="Cancel")]
                  ]

        self.window = sg.Window('Save...', layout)

    def progress_bar_layout(self):
        layout = [[sg.Text(f'Progressing {self.total_jpg} Images...', key='title')],
                  [sg.Text("0"), sg.ProgressBar(1, orientation='h', size=(20, 20), key='progress'),
                   sg.Text(str(self.total_jpg))],
                  ]
        self.window = sg.Window('Progress', layout)

    def create_progress_bar(self):
        self.window = self.window.Finalize()

    def progress_bar_update(self, cur_progress, total):
        self.window['progress'].UpdateBar(cur_progress, total)

    def close_progress_bar(self):
        self.window.Close()

    def create_window(self):
        self.event, self.values = self.window.read()
        self.window.close()
        if self.event in (sg.WIN_CLOSED, 'Cancel'):
            sys.exit()

    @staticmethod
    def no_insta_load_files():
        sg.popup_error(
            f"Couldn't find a json.xz-File in the given directory. Restart the tool and read the docs.",
            title='Files not found',
            keep_on_top=True
        )
        sys.exit()


if __name__ == '__main__':
    sg.popup_get_file(message="Hallo", title="Moin", multiple_files=True)
