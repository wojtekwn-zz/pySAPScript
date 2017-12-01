import re
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
    def __init__(self, connection):
        self.connection = connection
        self.script = []
        
    def __repr__(self):
        return '\n'.join(self.script)
        
    def loads(self,sapScriptFile):
        f = open(sapScriptFile, 'r+')
        self.__sap2python(f.readlines())
    
    def __processLine(self, line):
        line = line.replace('true', 'True').replace('false', 'False')
        cnt = re.split(r'(\s+)', line)
        
        if len(cnt) == 1:
            return cnt[0] + '()'
        else:
            if cnt[2][0] == '=':
                return line
            elif cnt[0] == '':
                return None
            else:
                return cnt[0] + '(' + "".join(cnt[2:]) + ')'
    
    def __sap2python(self, inputFileData):
        self.script = ['import win32com.client',
                       'import pythoncom',
                       'SapGui = win32com.client.GetObject("SAPGUI").GetScriptingEngine',
                       'try:',
                       '    session = SapGui.FindById("ses[0]")',
                       'except pythoncom.com_error:',
                       '    conn = SapGui.OpenConnection("' + self.connection + '", True)',
                       '    session = SapGui.FindById("ses[0]")']
        skip = False
        for line in inputFileData:
            if line[:2] == 'If':
                skip = True
            if not skip:
                self.script.append(self.__processLine(line[:-1]))
            if line[:6] == 'End If':
                skip = False
