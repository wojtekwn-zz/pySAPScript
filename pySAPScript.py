"""

This is a very basic approach for converting vba code generated with
SAP Scripting module of SAP GUI. There are three requiremants for the script to run:
 1) code has to be executed from the machine running SAP
    - citrix will not work
    - local instance of SAP with work just fine
 2) Windows machine
 3) win32com python library (pip install pywin32)
    - note that python and pywin32 can be installed on corporate machines
      without administrative rights, check anaconda project for details

Note that the script is very basic, but seems to work just fine with all of
our the processes in the company. Feedback is always welcome.

"""

class SAPScript:
    def __init__ (self, connection):
        self.connection = connection
        self.script = []
        
    def __repr__ (self):
        return '\n'.join(self.script)
        
    def loads(self,sapScriptFile):
        f = open(sapScriptFile, 'r+')
        self.__sap2python(f.readlines())
    
    def __processLine(self, line):
        cnt = line.split(' ')
        if len(cnt) == 1:
            return cnt[0] + '()'
        else:
            if cnt[1][0] == '=':
                return line
            else:
                return cnt[0] + '(' + cnt[1] + ')'
    
    def __sap2python(self, inputFileData):
        self.script = ['import win32com.client',
                     'SapGui = win32com.client.Dispatch("Sapgui.ScriptingCtrl.1")',
                     'conn = SapGui.OpenConnection("' + self.connection + '", True)',
                     'session = conn.Children[0]']
        skip = False
        for line in inputFileData:
            if line[:2] == 'If':
                skip = True
            if not skip:
                self.script.append(self.__processLine(line[:-1]))
            if line[:6] == 'End If':
                skip = False

