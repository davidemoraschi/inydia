# -*- coding: cp1252 -*-



import os.path
import win32api
import win32con


from iOrigen import *
from MyExceptions import * 


class GenAppOrigen(iOrigen):

    def ___init__(self):

        self.inputData   = ""
        self.count       = 0


    def start(self,inputData):

        try:
            if not inputData:                   raise ErrorInputOrigen
            
            self.inputData  = inputData

            tmp               = self.inputData.split(",")
            self.count        = (len(tmp)-1)/2

            if self.count <> int(tmp[0]):       raise ErrorInputOrigen

            k = 1

            self.testPaths   = []

            while k < len(tmp):

                self.testPaths.append(os.path.join(
                    os.path.splitdrive(self.leeRegistro("InstallDir"))[0] + os.sep,
                    "Genomica","Data","Results",
                    tmp[k],tmp[k+1]))
                k = k + 2


        except ErrorInputOrigen, e:

            raise ErrorInputOrigen(e)


        except Exception, e:

            raise ErrorAlSeleccionarCarpeta(e)


    def dameCarpeta(self):

        if self.testPaths:      return self.testPaths
        else:                   raise ErrorAlSeleccionarCarpeta


    def leeRegistro(self,variable):

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


                        
        
