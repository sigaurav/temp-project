import tkinter as tk
from tkinter import ttk, messagebox
from metadata import color_pallet
from tkinter import filedialog as fd
from main import SegregateFiles

filepath = ''

class OpenFileUI(tk.Tk):

    def __init__(self):
        super().__init__()

        self.title("Split Files")
        self.geometry('450x200')
        self.resizable(False, False)

        style = ttk.Style(self)
        style.theme_use('clam')

        self['background'] = color_pallet.get('COLOR_PRIMIARY')
        style.configure('Title.TLabel', font=('Lucida Sans', 16), background=color_pallet.get('COLOR_LIGHT_BG'),
                        foreground=color_pallet.get('COLOR_DARK_TEXT'))
        style.configure('TitleAnalysis.TLabel', font=('Lucida Sans', 12), background=color_pallet.get('COLOR_LIGHT_BG'),
                        foreground=color_pallet.get('COLOR_DARK_TEXT'))
        style.configure('Analysis.TLabel', font=('Lucida Sans', 10), background=color_pallet.get('COLOR_PRIMARY'),
                        foreground=color_pallet.get('COLOR_LIGHT'))
        style.configure('Body.TLabel', font=('helvetica', 12), background=color_pallet.get('COLOR_PRIMARY'),
                        foreground=color_pallet.get('COLOR_LIGHT'))
        style.configure('Background.TFrame', background=color_pallet.get('COLOR_LIGHT_BG'))
        style.configure('Analysis.TFrame', background=color_pallet.get('COLOR_PRIMARY'))
        style.configure('BackgroundButton.TFrame', background=color_pallet.get('COLOR_PRIMARY'))
        style.configure('BrowseButton.TButton', background=color_pallet.get('COLOR_SEC'),
                        foreground=color_pallet.get('COLOR_LIGHT_CONTRAST'))
        style.configure('BrowseButton.TButton', background=[('active', color_pallet.get('COLOR_SEC'))])
        style.configure('TNotebook.Tab', focuscolor=style.configure('.')['background'])
        style.configure('ElementComboBox.TCombobox', fieldbackgound=color_pallet.get('COLOR_LIGHT_CONTRAST'),
                        background=color_pallet.get('COLOR_LIGHT_CONTRAST'))
        style.configure('Definition.TCheckbutton', background=color_pallet.get('COLOR_PRIMARY'),
                        foreground=color_pallet.get('COLOR_LIGHT'), focuscolor=style.configure('.')['background'])
        style.configure('Definition.TCheckbutton', background=['active', (color_pallet.get('COLOR_PRIMARY'))],
                        foreground=['active', (color_pallet.get('COLOR_LIGHT'))])
        style.configure('Definition.TButton', background=color_pallet.get('COLOR_PRIMARY'),
                        foreground=color_pallet.get('COLOR_LIGHT'))
        style.configure('Definition.TButton', background=['active', (color_pallet.get('COLOR_PRIMARY'))],
                        foreground=['active', (color_pallet.get('COLOR_LIGHT'))])

        container = ttk.Frame(self)
        container.grid()

        self.title_frame = ttk.Frame(self, style='Background.TFrame')
        self.title_frame.grid(row=0, column=0)

        title = ttk.Label(self.title_frame, text='Split Files', style='Title.TLabel')
        title.grid(row=0, column=0, sticky='WE', padx=(160, 160), pady=(5, 5))

        self.body_frame = ttk.Frame(self, style='Analysis.TFrame')
        self.body_frame.grid(row=1, column=0, sticky='WE', padx=(10, 10), pady=(10, 10))

        load_file_button = ttk.Button(self.body_frame, text='Browse File', cursor='hand2', command=self.select_file)
        load_file_button.grid(row=2, column=0, sticky='WE', padx=(10, 10), pady=(20, 20))

        generate_files_button = ttk.Button(self.body_frame, text='Generate Files', cursor='hand2',
                                           command=self.generate_files)
        generate_files_button.grid(row=2, column=1, sticky='WE', padx=(10, 10), pady=(20, 20))

    def select_file(self):
        global filepath
        label_loading = ttk.Label(self.body_frame, text='Status: Loading . . .', style='Body.TLabel')
        label_loading.grid(row=2, column=2, sticky='WE', padx=(10, 10), pady=(10, 10))

        filetypes = (
            ('Excel', '.*xlsx'),
            ('All', '*.**')
        )

        filepath = fd.askopenfilename(
            title='Select File', filetypes=filetypes
        )

        label_loading = ttk.Label(self.body_frame, text='Status: File Loaded', style='Body.TLabel')
        label_loading.grid(row=2, column=2, sticky='WE', padx=(10, 10), pady=(10, 10))

    def generate_files(self):
        label_loading = ttk.Label(self.body_frame, text='Status: Creating Files . . .', style='Body.TLabel')
        label_loading.grid(row=2, column=2, sticky='WE', padx=(10, 10), pady=(10, 10))

        output_path = fd.askdirectory()
        SegregateFiles.process_masterfile(filepath, output_path)

        label_loading = ttk.Label(self.body_frame, text='Status: Successful', style='Body.TLabel')
        label_loading.grid(row=2, column=2, sticky='WE', padx=(10, 10), pady=(10, 10))
        exit(0)


# def main():
#     file = OpenFileUI()
#
#
# if __name__ == '__main__':
#     main()

file = OpenFileUI()
file.mainloop()