# -*- coding: cp1252 -*-


from Diccionario import *
from iExportador import *
from ExportThread import *
from ConfigLoader import *
from ProgressDialog import *
from FinishWindow import *
import logging




class GenAppExportador(iExportador):


    def __init__(self):

        self.myConfig = ConfigLoader("GenAppConfig.xml")


    def start(self,inputData):

        self.thread        = ExportThread(inputData,self.myConfig)

        self.myProgressWin = ProgressDialog(self.thread)
        self.thread.notifyToWindow(self.myProgressWin)

        self.myFinishWindow = FinishWindow()

        
    def exporta(self):

        self.myProgressWin.Show()
        self.thread.join()

        self.myFinishWindow.Show(self.thread.getProgress())

        self.myProgressWin.Unregister()
        self.myFinishWindow.Unregister()


            
