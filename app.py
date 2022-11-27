import tkinter as tk
from tkinter import messagebox, filedialog as fd
import tkinter.ttk as ttk
from main import SegregateFiles
from main import MergeFiles
from PIL import Image, ImageTk


class OpenUI():

    filePath = ''
    def __init__(self, master=None):
        # build ui
        self.style = ttk.Style(root)
        self.Frame = ttk.Frame(master)
        self.Frame.configure(height=200, width=500)

        # Select File Button
        #############
        self.bg = ImageTk.PhotoImage(Image.open('asset/horses.png').resize((400, 70), Image.Resampling.LANCZOS))
        self.canvas1 = tk.Canvas(self.Frame, height=200, width=500)
        self.canvas1.place(anchor='nw')
        self.canvas1.create_image((10, 99), anchor='nw', image=self.bg)
        self.Frame.pack(side='top')
        #############

        self.browseButton = ttk.Button(self.Frame, text='Select File', cursor='hand2',
                                       command=self.select_file)
        self.browseButton.place(
            anchor="nw",
            relx=0.18,
            rely=0.36,
            width=330,
            x=0,
            y=0)

        # Split Button
        self.splitButton = ttk.Button(self.Frame, cursor='hand2', state='disabled', text='Split Files', command=self.generate_files)
        self.splitButton.place(
            anchor="nw",
            relx=0.18,
            rely=0.53,
            width=158,
            x=0,
            y=0)

        # Merge Button
        self.mergeButton = ttk.Button(self.Frame, cursor="hand2", text='Merge Files', command=self.merge_files)
        self.mergeButton.place(
            anchor="nw",
            relx=0.523,
            rely=0.53,
            width=158,
            x=0,
            y=0)

        # Label
        self.label = ttk.Label(self.Frame)
        self.label.configure(
            background="#c0c0c0",
            font="{Calibri} 20 {}",
            foreground="#800040",
            justify="left",
            text='Welcome To SMF')

        self.label.place(anchor="nw", relx=0.33, rely=0.05, x=0, y=0)
        # Tag Line
        self.tagLine = ttk.Label(self.Frame)
        self.tagLine.configure(
            font="{Cascadia Code} 8 {}",
            text='Making Split & Merge Interesting')
        self.tagLine.place(anchor="nw", relx=0.326, rely=0.24, x=0, y=0)

        # Canvas
        canvas2 = tk.Canvas(self.Frame)
        canvas2.configure(height=60, relief="ridge", width=70)
        canvas2.place(anchor="nw", relx=0.05, rely=0.045, x=0, y=0)
        self.img = ImageTk.PhotoImage(Image.open('asset/wellsLogo.png').resize((90, 90), Image.Resampling.LANCZOS))
            # self.img = tk.PhotoImage(file='asset/eAR Logo.jpeg')
        canvas2.create_image((0, 0), anchor='nw', image=self.img)
        self.Frame.pack(side="top")

        # Progress Bar
        self.style.layout("LabeledProgressbar",
         [('LabeledProgressbar.trough',
           {'children': [('LabeledProgressbar.pbar',
                          {'side': 'left', 'sticky': 'ns'}),
                         ("LabeledProgressbar.label",   # label inside the bar
                          {"sticky": ""})],
           'sticky': 'nswe'})])
        self.progressBar = ttk.Progressbar(self.Frame, orient='horizontal', style='LabeledProgressbar', cursor='wait', mode='determinate')
        self.progressBar.place(anchor="nw", width=500, x=0, y=182)

        # Main widget
        self.mainwindow = self.Frame

    def run(self):
        self.mainwindow.mainloop()

    def self(self):
        pass

    def merge_files(self):

        filez = fd.askopenfilenames(title='Select files to Merge', filetypes=(('Excel','.xlsx'), ('All', '*.*')))
        if len(filez) == 0 or filez is None:
            messagebox.showerror('Missing Path', 'No Directory Selected')
            exit(0)
        try:
            output_path = fd.askdirectory(title='Select Output Path')
            MergeFiles.merge_files(filez, output_path, self.progressBar)
        except Exception as e:
            self.style.configure('LabeledProgressbar', text='Status: Failed...')
            tk.messagebox.showerror('Failed', e)

        self.style.configure('LabeledProgressbar', text='Status: Successful...')
        messagebox.showinfo('Status:', 'Process Completed')
        exit(0)


    def select_file(self):
        global filepath
        self.style.configure('LabeledProgressbar', text='Selecting File...')

        filetypes = (
            ('Excel', '.*xlsx'),
            ('All', '*.**')
        )

        filepath = fd.askopenfilename(
            title='Select File', filetypes=filetypes
        )

        if filepath == '' or filepath is None:
            self.style.configure('LabeledProgressbar', text='!!No File Selected!!')
        else:
            self.style.configure('LabeledProgressbar', text='!!File Loaded!!')
            self.splitButton.configure(state='enabled')

    def generate_files(self):

        if filepath == '' or filepath is None:
            messagebox.showerror('Missing File', 'No Input found')
            exit(0)
        else:
            self.style.configure('LabeledProgressbar', text='Creating Files...')

        output_path = fd.askdirectory(title='Select Output Path')
        if output_path == '' or output_path is None:
            messagebox.showerror('Missing Path', 'No Output Path Given')
            exit(0)
        try:
            SegregateFiles.process_masterfile(filepath, output_path, self.progressBar)
        except Exception as e:
            self.style.configure('LabeledProgressbar', text='Status: !!!Failed!!')
            tk.messagebox.showerror('Failed', e)

        self.style.configure('LabeledProgressbar', text='Status: Successful')
        messagebox.showinfo('Status:', 'Process Completed')
        exit(0)


if __name__ == "__main__":
    root = tk.Tk()
    root.title('Split/Merge Files')
    app = OpenUI(root)
    app.run()
