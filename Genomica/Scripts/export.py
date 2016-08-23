# -*- coding: cp1252 -*-


import logging
import os
from subprocess import *
import win32api
import win32con



def LeeRegistro(variable):

        try:
                keyHandle = win32api.RegOpenKeyEx(win32con.HKEY_LOCAL_MACHINE,"Software\\Genomica",0,win32con.KEY_ALL_ACCESS)

                try:
                        value,typeId = win32api.RegQueryValueEx(keyHandle,variable)

                except:
                        value,typeId = win32api.RegQueryValueEx(keyHandle,variable)

                win32api.RegCloseKey(keyHandle)

        except:
                value = ""

        return value



def main(inputOrigen = "",modulesFile = "",mockArgument1 = "",mockArgument2 = ""):

    try:

        pythonExe = "C:\\python25\\python.exe"
        exportPy  = os.path.join(
            os.path.splitdrive(LeeRegistro("InstallDir"))[0] + os.sep,
            "genomica","data","scripts","exporter.py")

        command = pythonExe + " " + exportPy + " " + \
                  inputOrigen + " " + \
                  modulesFile + " " + \
                  mockArgument1 + " " + \
                  mockArgument2
        
        call(command,shell=True)
        
        return 1


    except Exception, e:

        return 0



if __name__ == "__main__":



    try:
        argv1 = sys.argv[1]

    except:
        argv1 = ""

    try:
        argv2 = sys.argv[2]

    except:
        argv2 = ""

    try:
        argv3 = sys.argv[3]

    except:
        argv3 = ""

    try:
        argv4 = sys.argv[4]

    except:
        argv4 = ""

    
    main(argv1,argv2,argv3,argv4)


