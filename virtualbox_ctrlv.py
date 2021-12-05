from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
from tkinter import ttk
import subprocess
import re
import winreg
import pathlib
import sys

class Application(ttk.Frame):
    ComboBox1 = False
    Button1 = False
    Button2 = False
    def __init__(self, master=None):
        super().__init__(master, padding='3 3 12 12')
        self.master = master
        self.grid(column=0, row=0, sticky=(N, W, E, S))
        self.getpath()
        self.getvms()
        self.create_widgets()

    def getpath(self):
        # HKEY_LOCAL_MACHINE\SOFTWARE\Oracle\VirtualBox
        # InstallDir
        # C:\Program Files\Oracle\VirtualBox\
        default_oracle_path = 'C:\\Program Files\\Oracle\\VirtualBox\\'
        oracle_path = False
        regkey = False
        foundinregistry = True
        try:
            regkey = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 'SOFTWARE\\Oracle\\VirtualBox')
        except FileNotFoundError:
            import platform
            bitness = platform.architecture()[0]
            if bitness == '32bit':
                other_view_flag = winreg.KEY_WOW64_64KEY
            elif bitness == '64bit':
                other_view_flag = winreg.KEY_WOW64_32KEY
            try:
                regkey = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 'SOFTWARE\\Oracle\\VirtualBox', access=winreg.KEY_READ | other_view_flag)
            except FileNotFoundError:
                foundinregistry = False
        if foundinregistry:
            try:
                oracle_path = winreg.QueryValueEx(regkey, 'InstallDir')[0]
                if oracle_path[-1:] != '\\':
                    oracle_path += '\\'
            except FileNotFoundError:
                foundinregistry = False
        self.vbm_path = False
        if not foundinregistry or not pathlib.Path(oracle_path + 'VBoxManage.exe').is_file():
            oracle_path = default_oracle_path
            if not pathlib.Path(oracle_path + 'VBoxManage.exe').is_file():
                if messagebox.askyesno('Cannot find installed VirtualBox', 'Will you find VBoxManage.exe manually?'):
                    self.vbm_path = filedialog.askopenfilename(filetypes=[('VBoxManage.exe','VBoxManage.exe')])
                if not self.vbm_path or not pathlib.Path(vbm_path).is_file():
                    self.master.destroy()
                    sys.exit()
        if not self.vbm_path:
            self.vbm_path = oracle_path + 'VBoxManage.exe'

    def populate_combobox(self):
        if self.ComboBox1:
            if self.vms:
                self.ComboBox1['values'] = [x[0] for x in self.vms]
                self.ComboBox1.current(0)
                self.ComboBox1.state(['!disabled'])
                if self.Button2:
                    self.Button2.state(['!disabled'])
            else:
                self.ComboBox1.current(-1)
                self.ComboBox1['values'] = []

    def getvms(self):
        if self.Button1:
            self.Button1.state(['disabled'])
        if self.Button2:
            self.Button2.state(['disabled'])
        if self.ComboBox1:
            self.ComboBox1.state(['disabled'])
        cp = subprocess.run('"' + self.vbm_path + '" list runningvms', capture_output=True, encoding='mbcs')
        self.returncode = cp.returncode
        self.stderr = cp.stderr
        self.vms = []
        if cp.returncode == 0:
            for cpo in cp.stdout.splitlines():
                m = re.match('"(.*)" \{([^\}]+)\}', cpo)
                self.vms.append((m[1], m[2]))
        self.populate_combobox()
        if self.Button1:
            self.Button1.state(['!disabled'])

    def str_to_scancodes(self, s):
        unshifted = {
            '\x1b': 0x01, '1': 0x02, '2': 0x03, '3': 0x04, '4': 0x05, '5': 0x06,
            '6': 0x07, '7': 0x08, '8': 0x09, '9': 0x0a, '0': 0x0b, '-': 0x0c,
            '=': 0x0d, '\b': 0x0e, '\t': 0x0f, 'q': 0x10, 'w': 0x11, 'e': 0x12,
            'r': 0x13, 't': 0x14, 'y': 0x15, 'u': 0x16, 'i': 0x17, 'o': 0x18,
            'p': 0x19, '[': 0x1a, ']': 0x1b, '\n': 0x1c, 'a': 0x1e, 's': 0x1f,
            'd': 0x20, 'f': 0x21, 'g': 0x22, 'h': 0x23, 'j': 0x24, 'k': 0x25,
            'l': 0x26, ';': 0x27, "'": 0x28, '`': 0x29, '\\': 0x2b, 'z': 0x2c,
            'x': 0x2d, 'c': 0x2e, 'v': 0x2f, 'b': 0x30, 'n': 0x31, 'm': 0x32,
            ',': 0x33, '.': 0x34, '/': 0x35, ' ': 0x39,
        }
        shifted = {
            '!': 0x02, '@': 0x03, '#': 0x04, '$': 0x05, '%': 0x06, '^': 0x07,
            '&': 0x08, '*': 0x09, '(': 0x0a, ')': 0x0b, '_': 0x0c, '+': 0x0d,
            'Q': 0x10, 'W': 0x11, 'E': 0x12, 'R': 0x13, 'T': 0x14, 'Y': 0x15,
            'U': 0x16, 'I': 0x17, 'O': 0x18, 'P': 0x19, '{': 0x1a, '}': 0x1b,
            'A': 0x1e, 'S': 0x1f, 'D': 0x20, 'F': 0x21, 'G': 0x22, 'H': 0x23,
            'J': 0x24, 'K': 0x25, 'L': 0x26, ':': 0x27, '"': 0x28, '~': 0x29,
            '|': 0x2b, 'Z': 0x2c, 'X': 0x2d, 'C': 0x2e, 'V': 0x2f, 'B': 0x30,
            'N': 0x31, 'M': 0x32, '<': 0x33, '>': 0x34, '?': 0x35,
        }
        shiftstate = False
        ctrlstate = False
        out = []
        for c in s:
            if c in unshifted and not (ctrlstate and c in ['\x1b', '\b', '\t', '\n']):
                if ctrlstate:
                    out.append('9d')
                    ctrlstate = False
                if shiftstate:
                    out.append('aa')
                    shiftstate = False
                out.extend(['%02x' % unshifted[c], '%02x' % (unshifted[c] + 0x80)])
            elif c in shifted:
                if ctrlstate:
                    out.append('9d')
                    ctrlstate = False
                if not shiftstate:
                    out.append('2a')
                    shiftstate = True
                out.extend(['%02x' % shifted[c], '%02x' % (shifted[c] + 0x80)])
            elif ord(c) in range(32) or c == '\x7f':
                if shiftstate:
                    out.append('aa')
                    shiftstate = False
                if not ctrlstate:
                    out.append('1d')
                    ctrlstate = True
                if c == '\x7f':
                    out.extend(['35', 'b5'])
                else:
                    b = shifted[chr(ord(c) + 0x40)]
                    out.extend(['%02x' % b, '%02x' % (b + 0x80)])
            else:
                messagebox.showerror('Error', 'Character "' + c + '" (' + hex(ord(c)) + ') is outside of allowed range')
        if ctrlstate:
            out.append('9d')
            ctrlstate = False
        if shiftstate:
            out.append('aa')
            shiftstate = False
        return ' '.join(out)

    def sendtext(self):
        if self.ComboBox1.current() == -1:
            messagebox.showerror('Error', 'No VM is selected')
            return
        self.Button1.state(['disabled'])
        self.Button2.state(['disabled'])
        self.ComboBox1.state(['disabled'])
        cp = subprocess.run(
            '"' + self.vbm_path + '" controlvm ' + self.vms[self.ComboBox1.current()][1] + ' keyboardputscancode ' + self.str_to_scancodes(self.TextBox1.get('1.0', 'end')[:-1]),
            capture_output=True,
            encoding='mbcs'
        )
        self.Button1.state(['!disabled'])
        self.Button2.state(['!disabled'])
        self.ComboBox1.state(['!disabled'])

    def combobox_selected(self, dummy):
        self.ComboBox1.selection_clear()

    def create_widgets(self):
        s = ttk.Style()
        s.configure('Emergency.TButton', font='helvetica 24', foreground='red', padding=10)
        self.ComboBox1 = ttk.Combobox(self)
        self.ComboBox1.state(['readonly'])
        self.ComboBox1.bind('<<ComboboxSelected>>', self.combobox_selected)
        self.populate_combobox()
        self.ComboBox1.grid(column=0, row=0, sticky=W)
        self.Button1 = ttk.Button(self, text='Resresh', command=self.getvms)
        self.Button1.grid(column=1, row=0, sticky=W)
        self.TextBox1 = Text(self, width=40, height=10)
        self.TextBox1.grid(column=0, row=1, columnspan=2, sticky=W)

        self.Button2 = ttk.Button(self, text='Send', style='Emergency.TButton', command=self.sendtext)
        self.Button2.grid(column=0, row=2, sticky=W)
        self.master.state('normal')

root = Tk()
root.withdraw()
app = Application(master=root)
app.mainloop()
