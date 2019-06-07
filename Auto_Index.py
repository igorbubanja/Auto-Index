import os
import csv
from cmd import Cmd
import shutil
import pickle

import tkinter as tk
from tkinter import filedialog
from tkinter.filedialog import askopenfilename

'''TODO:
    Error handling for not having a dms.csv present, and not having 'copied' and 'numbered' folders present
    '''

copy_location = ''
indexed_location = ''
dms_location = ''

FILENAME_INDEX = 12 #usually 12, changed for Nebi

class Drawing:
 
    def __init__(self, num, sht, rev, index, filename):
        self.num = num
        self.sht = sht
        self.rev = rev
        self.index = index
        self.filename = filename
        
def select_directory(): #used to select location of numbered and copied folders
    filepath = ''
    root = tk.Tk()
    root.withdraw() #use to hide tkinter window
    
    currdir = os.getcwd()
    tempdir = filedialog.askdirectory(parent=root, initialdir=currdir, title='Please select a directory')
    if len(tempdir) > 0:
        filepath = tempdir
    return filepath

def select_file(): #used to select dms location
    tk.Tk().withdraw()
    filename = askopenfilename() 
    if (len(filename)>0):
        return filename
    
def set_paths_from_memory():
    global copy_location
    global indexed_location
    global dms_location
    
    try:
        with open('copy.pickle', 'rb') as f:
            copy_location = pickle.load(f)
            f.close()
            print('success')
    except FileNotFoundError:
        copy_location
    
        
    try:
        with open('index.pickle', 'rb') as f:
            indexed_location = pickle.load(f)
            f.close()
    except FileNotFoundError:
        indexed_location = ''
        
    try:
        with open('dms.pickle', 'rb') as f:
            dms_location = pickle.load(f)
            f.close()
    except FileNotFoundError:
        dms_location = ''
    

def index_value_pairs():
    my_dict = {}
    with open(dms_location, 'r') as f: #Mode should be 'rb' in python 2 and 'r' in python 3
        reader = csv.reader(f)
        for row in reader:
            index = row[1]
            file_name = row[FILENAME_INDEX].lower()
            if(len(index)==0):
            
                continue
            if(index[-1] == 'H'):
                continue
            my_dict[file_name] = index
    return my_dict

def drawing_list():
    drawing_list = []
    try:
        with open(dms_location, 'r') as f: #Mode should be 'rb' in python 2 and 'r' in python 3
    
            reader = csv.reader(f)
            for row in reader:
                index = row[1]
                drawing_number = row[14]
                sheet = row[15]
                revision = row[16]
                filename = row[FILENAME_INDEX]
                
                drawing = Drawing(drawing_number, sheet, revision, index, filename)
                drawing_list.append(drawing)
            
    except FileNotFoundError:
        print('Make sure you have selected the correct DMS Location')
				
    return drawing_list

def addRev():
    string_to_replace = input('What string should be replaced with the revision (e.g. Model): ')
    filepath = copy_location
    drw_list = drawing_list()
    for filename in os.listdir(filepath):
        filename_without_model = filename.replace(string_to_replace + '.pdf' , '')
        
        for d in drw_list:
            if filename_without_model.lower() in d.filename.replace(d.rev, '').lower() and d.index.isnumeric():

                new_filename = filename.replace(string_to_replace, d.rev)
                try:
                    os.rename(os.path.join(filepath, filename), os.path.join(filepath, new_filename))
                except FileNotFoundError:
                    print('File not found')
                except FileExistsError:
                    os.rename(os.path.join(filepath, filename), os.path.join(filepath, 'Already renamed-' + filename))
    print('done')

def rename3():
    drw_list = drawing_list()
    my_dict = index_value_pairs()
    filepath = copy_location
    renamed_files = []
    for filename in os.listdir(filepath):
        if filename.lower()[-3:] != 'pdf': #skip files that aren't PDFs
            continue
        if filename.lower()[:-4] in [d.filename.lower() for d in drw_list]:
            try:
                print(filename)
                print(my_dict[filename.lower()[:-4]])
                os.rename(os.path.join(filepath, filename), os.path.join(indexed_location, my_dict[filename.lower()[:-4]] + ' ' + filename))
            except KeyError:
                print('Not found ' + filename.lower()[:-4] + '. Check that an index number is allocated to this drawing in the DMS')
            except FileExistsError:
                    os.rename(os.path.join(filepath, filename), os.path.join(filepath, 'Already renamed-' + filename))
        renamed_files.append(filename)
    print('complete')
        
def setcopylocation():
        global copy_location
        copy_location = select_directory()
        with open('copy.pickle', 'wb') as f:  
            pickle.dump(copy_location, f)
            f.close()
        
def getcopylocation():
    global copy_location
    print(copy_location)

def setindexlocation():
    global indexed_location
    indexed_location = select_directory()
    with open('index.pickle', 'wb') as f:  
        pickle.dump(indexed_location, f)
        f.close()

def getindexlocation():
    global indexed_location
    print(indexed_location)
    
def set_dms():
    global dms_location
    dms_location = select_file()
    with open('dms.pickle', 'wb') as f:  
        pickle.dump(dms_location, f)
        f.close()
    
def get_dms():
    global dms_location
    print(dms_location)
                    
class prompt(Cmd):
    def do_hello(self, args):
        print("hello: " + args)
		
    def do_rename3(self, args):
        '''Function that adds the drawing index to the beginning of the PDF filename'''
        rename3()
        
    def do_add1A(self, args):
        '''function doesn't work well in its current state. Currently is just adds 1A to the end of everything in the copy location. I did this to add a string at the end of files so that I can run addRev on file with a filename ending with just 2 dashes'''
        #filepath = r'''copied-1a'''
        filepath = copy_location
        renamed_files = []
        for filename in os.listdir(filepath):
            if filename.lower()[-3:] != 'pdf': #skip files that aren't PDFs
                continue
            os.rename(os.path.join(filepath, filename), os.path.join(filepath, (filename[:-4] + '-1A.pdf')))
            
    def do_addRev(self, args):
        '''Renames files with 'Model', 'Layout', etc in the filename to having the revision in the filename, just use this once for the cancelled drawings, it isn't written well'''
        addRev()
       
                    
    def do_copyall(self, args):
        dest = r'''copied'''
        for r, _, files in os.walk(r'''CANCELLED & SS DWGS'''):
            for filename in files:
                if filename.lower()[-3:] != 'pdf':
                    continue
                shutil.copy(os.path.join(r, filename), dest)
        print('done')
        
    def do_setcopylocation(self, args):
        setcopylocation()
        
    def do_getcopylocation(self, args):
        getcopylocation()
    
    def do_setindexlocation(self, args):
        setindexlocation()
    
    def do_getindexlocation(self, args):
        getindexlocation()
        
    def do_set_dms(self, args):
        set_dms()
        
    def do_get_dms(self, args):
        get_dms()
        
    #testing functions
    
    def do_dwg(self, args):
        for d in drawing_list():
            print(d.index)
            
    def do_files(self, args):
        filepath = copy_location
        for filename in os.listdir(filepath):
            print(filename)
            
    def do_fail(self, args):
        '''purposely fail as it is quicker than quitting the console'''
        b = int(s)
        
    def do_undo(self, args):
        global indexedlocation
        filepath = indexed_location
        for filename in os.listdir(filepath):
            space_index = 0
            for i in range(len(filename)):
                if filename[i] is ' ':
                    space_index = i
                    print(space_index)
            os.rename(os.path.join(filepath, filename), os.path.join(filepath, filename[space_index+1:]))
        print('complete')

        
'''Functions for GUI'''

    
#def set_paths_gui():
#    global dms_location
#    global indexed_location
#    global copy_location
#    new_window = tkinter.Toplevel(window)
#    dms_label = tkinter.Label(new_window, text = 'test', fg = 'red')
#    dms_label.grid(row = 0, column = 1, sticky = 'ew')
#    tkinter.Button(new_window, text = 'Set DMS location', command = set_dms_button).grid(row = 0, sticky = 'ew')
#    copy_label = tkinter.Label(new_window, text = copy_location).grid(row = 0, column = 1, sticky = 'ew')
#    tkinter.Button(new_window, text = 'Set Copy location', command = setcopylocation).grid(row = 1, sticky = 'ew')
#    new_window.mainloop()
    
def ifNotNone(x):
    if x is not None and x is not '':
        return x
    else:
        return 'Not Selected'

class MainWindow(tk.Frame):
    counter = 0
    def __init__(self, *args, **kwargs):
        tk.Frame.__init__(self, *args, **kwargs)
        self.button1 = tk.Button(self, text="Set Paths", 
                                command=self.set_paths)
        self.button2 = tk.Button(self, text = 'Execute',
                                command = self.execute)
        self.button1.pack(side='top')
        self.button2.pack(side ='top')
        
        
    def execute(self):
        rename3()
        
    def set_paths(self):
        global dms_location
        global copy_location
        global indexed_location
        dmslabel = ifNotNone(dms_location)
        copylabel = ifNotNone(copy_location)
        indexlabel = ifNotNone(indexed_location)
        self.t = tk.Toplevel(self)
        self.t.wm_title('Set File Paths')
        self.dms_label = tk.Label(self.t, text = dmslabel)
        self.copy_label = tk.Label(self.t, text = copylabel)
        self.index_label = tk.Label(self.t, text = indexlabel)
        dms_button = tk.Button(self.t, text = 'Set DMS location', command = self.set_dms_button)
        copy_button = tk.Button(self.t, text = 'Set Copied location', command = self.set_copy_button)
        index_button = tk.Button(self.t, text = 'Set Indexed location', command = self.set_index_button)
        
        
        self.dms_label.grid(row = 0, column = 1, sticky = 'ew')
        dms_button.grid(row = 0, column = 0, sticky = 'ew')
        self.copy_label.grid(row = 1, column = 1, sticky = 'ew')
        copy_button.grid(row = 1, column = 0, sticky = 'ew')
        self.index_label.grid(row = 2, column = 1, sticky = 'ew')
        index_button.grid(row = 2, column = 0, sticky = 'ew')
        
    def set_dms_button(self):
        global dms_location
        set_dms()
        dmslabel = ifNotNone(dms_location)
        self.dms_label.config(text = dmslabel)
        
        
    def set_copy_button(self):
        global copy_location
        setcopylocation()
        copylabel = ifNotNone(copy_location)
        self.copy_label.config(text = copylabel)
    
    def set_index_button(self):
        global indexed_location
        setindexlocation()
        indexlabel = ifNotNone(indexed_location)
        self.index_label.config(text = indexlabel)
     
if __name__ == '__main__':
    set_paths_from_memory()
    mode = input('gui or cmd?: ')
    if (mode == 'cmd'):
        set_paths_from_memory()
        my_prompt = prompt()
        my_prompt.prompt = '>'
        my_prompt.cmdloop('--------Script to rename the PDFs in the folder---\n'\
                          '-------Type "help" for a list of commands-----------\n'\
                          '---------------Written by Igor--------------------\n')
    elif (mode == 'gui'):
        root = tk.Tk()
        main = MainWindow(root)
        main.pack(side="top", fill="both", expand=True)
        root.mainloop()
        
    else:
        print("Invalid input, please type either 'gui' or 'cmd'")
            
    
