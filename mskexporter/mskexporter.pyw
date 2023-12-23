"""A simple tkinter-based application for Manuskript data export from outsides.

Copyright (c) 2023 Peter Triesberger
For further information see https://github.com/peter88213/manuskript_md
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import os
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox

import mskmd
import pypandoc
import tkinter as tk

OUTPUT_TYPES = [
    'md',
    'odt',
    'docx',
    'html',
    ]


class MainTk:
    """A tkinter GUI root class."""
    _KEY_RESTORE_STATUS = ('<Escape>', 'Esc')
    _KEY_OPEN_PROJECT = ('<Control-o>', 'Ctrl-O')
    _KEY_QUIT_PROGRAM = ('<Control-q>', 'Ctrl-Q')

    def __init__(self, title, **kwargs):
        """Initialize the GUI window and instance variables.
        
        Positional arguments:
            title -- application title to be displayed at the window frame.
         
        Processed keyword arguments:
            root_geometry: str -- geometry of the root window.
        
        Operation:
        - Create a main menu to be extended by subclasses.
        - Create a title bar for the project title.
        - Open a main window frame to be used by subclasses.
        - Create a status bar to be used by subclasses.
        - Create a path bar for the project file path.
        
        Extends the superclass constructor.
        """
        self.usePandoc = True
        self.infoHowText = ''
        self._fileTypes = [('Manuskript project', '.msk')]
        self.title = title
        self._statusText = ''
        self.kwargs = kwargs
        self.prjDir = None
        self.root = tk.Tk()
        self.root.protocol("WM_DELETE_WINDOW", self.on_quit)
        self.root.title(title)
        if kwargs.get('root_geometry', None):
            self.root.geometry(kwargs['root_geometry'])
        self.mainMenu = tk.Menu(self.root)

        self._build_main_menu()
        # Hook for subclasses

        self.root.config(menu=self.mainMenu)
        self.mainWindow = tk.Frame()
        self.mainWindow.pack(expand=True, fill='both')
        self.statusBar = tk.Label(self.root, text='', anchor='w', padx=5, pady=2)
        self.statusBar.pack(expand=False, fill='both')
        self.statusBar.bind('<Button-1>', self.restore_status)
        self.pathBar = tk.Label(self.root, text='', anchor='w', padx=5, pady=3)
        self.pathBar.pack(expand=False, fill='both')

        #--- Output selector.
        self.outputType = tk.IntVar(value=0)
        for i, outputType in enumerate(OUTPUT_TYPES):
            ttk.Radiobutton(
                self.mainWindow,
                variable=self.outputType,
                value=i,
                text=outputType,
                ).pack(anchor='w', padx=10)

        #--- Event bindings.
        self.root.bind(self._KEY_RESTORE_STATUS[0], self.restore_status)
        self.root.bind(self._KEY_OPEN_PROJECT[0], self._open_project)
        self.root.bind(self._KEY_QUIT_PROGRAM[0], self.on_quit)

    def call_pandoc(self, fileList):
        """Call pandoc to convert Markdown files to s specified output format.
        
        Positional arguments:
            fileList: str -- List of paths to markdown files to be converted.
        """
        outputIndex = self.outputType.get()
        if outputIndex > 0:
            extension = OUTPUT_TYPES[outputIndex]
            for mdFile in fileList:
                dir, file = os.path.split(mdFile)
                name, __ = os.path.splitext(file)

                try:
                    pypandoc.convert_file(
                        mdFile,
                        extension,
                        outputfile=os.path.join(dir, f'{name}.{extension}'),
                        format='markdown-smart'
                        )
                except Exception as ex:
                    self.set_info_how(str(ex))

    def close_project(self, event=None):
        """Close the yWriter project without saving and reset the user interface.
        
        To be extended by subclasses.
        """
        self.prjDir = None
        self.root.title(self.title)
        self.show_status('')
        self.show_path('')
        self.disable_menu()

    def disable_menu(self):
        """Disable menu entries when no project is open.
        
        To be extended by subclasses.
        """
        self.fileMenu.entryconfig('Close', state='disabled')
        self.mainMenu.entryconfig('Convert', state='disabled')

    def enable_menu(self):
        """Enable menu entries when a project is open.
        
        To be extended by subclasses.
        """
        self.fileMenu.entryconfig('Close', state='normal')
        self.mainMenu.entryconfig('Convert', state='normal')

    def on_quit(self, event=None):
        """Save keyword arguments before exiting the program."""
        self.kwargs['root_geometry'] = self.root.winfo_geometry()
        self.root.quit()

    def open_project(self, fileName):
        """Create a yWriter project instance and read the file.

        Positional arguments:
            fileName: str -- project file path.
            
        Display project title and file path.
        Return True on success, otherwise return False.
        To be extended by subclasses.
        """
        self.restore_status()
        fileName = self.select_project(fileName)
        if not fileName:
            return False

        if self.prjDir is not None:
            self.close_project()
        dir, file = os.path.split(fileName)
        self.prjName, extension = os.path.splitext(file)
        self.prjDir = os.path.join(dir, self.prjName)
        if not os.path.isdir(self.prjDir):
            self.prjDir = None
            self.prjName = None
            return False

        self.show_path(f'{os.path.normpath(self.prjDir)}')
        self.enable_menu()
        return True

    def restore_status(self, event=None):
        """Overwrite error message with the status before."""
        self.show_status(self._statusText)

    def select_project(self, fileName):
        """Return a project file path.

        Positional arguments:
            fileName: str -- project file path.
            
        Optional arguments:
            fileTypes -- list of tuples for file selection (display text, extension).

        Priority:
        1. use file name argument
        2. open file select dialog

        On error, return an empty string.
        """
        initDir = os.path.dirname(self.kwargs.get('yw_last_open', ''))
        if not initDir:
            initDir = './'
        if not fileName or not os.path.isfile(fileName):
            fileName = filedialog.askopenfilename(filetypes=self._fileTypes, defaultextension='.yw7', initialdir=initDir)
        if not fileName:
            return ''

        return fileName

    def set_info_how(self, message):
        """Show how the converter is doing.
        
        Positional arguments:
            message -- message to be displayed. 
            
        Display the message at the status bar.
        Overrides the superclass method.
        """
        if message.startswith('!'):
            self.statusBar.config(bg='red')
            self.statusBar.config(fg='white')
            self.infoHowText = message.split('!', maxsplit=1)[1].strip()
        else:
            self.statusBar.config(bg='green')
            self.statusBar.config(fg='white')
            self.infoHowText = message
        self.statusBar.config(text=self.infoHowText)

    def show_path(self, message):
        """Put text on the path bar."""
        self._pathText = message
        self.pathBar.config(text=message)

    def show_status(self, message):
        """Put text on the status bar."""
        self._statusText = message
        self.statusBar.config(bg=self.root.cget('background'))
        self.statusBar.config(fg='black')
        self.statusBar.config(text=message)

    def start(self):
        """Start the Tk main loop.
        
        Note: This can not be done in the constructor method.
        """
        self.root.mainloop()

    def _build_main_menu(self):
        """Add main menu entries.
        
        This is a template method that can be overridden by subclasses. 
        """
        self.fileMenu = tk.Menu(self.mainMenu, tearoff=0)
        self.mainMenu.add_cascade(label='File', menu=self.fileMenu)
        self.fileMenu.add_command(label='Open...', accelerator=self._KEY_OPEN_PROJECT[1], command=lambda: self.open_project(''))
        self.fileMenu.add_command(label='Close', command=self.close_project)
        self.fileMenu.entryconfig('Close', state='disabled')
        self.fileMenu.add_command(label='Exit', accelerator=self._KEY_QUIT_PROGRAM[1], command=self.on_quit)

        self.cnvMenu = tk.Menu(self.mainMenu, tearoff=0)
        self.cnvMenu.add_command(label='Outline', command=self.convert_outline)
        self.cnvMenu.add_command(label='World', command=self.convert_world)
        self.cnvMenu.add_command(label='Characters', command=self.convert_characters)
        self.mainMenu.add_cascade(label='Convert', menu=self.cnvMenu)

        self.disable_menu()

    def _open_project(self, event=None):
        """Create a yWriter project instance and read the file.
        
        This non-public method is meant for event handling.
        """
        self.open_project('')

    def convert_outline(self, event=None):
        try:
            fileList = mskmd.convert_outline(self.prjDir)
            self.set_info_how('Outline documents written.')
        except Exception as ex:
            self.set_info_how(f'!{str(ex)}')
        else:
            self.call_pandoc(fileList)

    def convert_world(self, event=None):
        try:
            fileList = mskmd.convert_world(self.prjDir)
            self.set_info_how('Story world document written.')
        except Exception as ex:
            self.set_info_how(f'!{str(ex)}')
        else:
            self.call_pandoc(fileList)

    def convert_characters(self, event=None):
        try:
            fileList = mskmd.convert_characters(self.prjDir)
            self.set_info_how('Character sheets document written.')
        except Exception as ex:
            self.set_info_how(f'!{str(ex)}')
        else:
            self.call_pandoc(fileList)


if __name__ == '__main__':
    app = MainTk('Manuskript to Markdown converter', root_geometry='400x200')
    app.start()
