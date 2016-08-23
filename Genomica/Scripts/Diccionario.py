# -*- coding: cp1252 -*-

import win32api
import win32con
import logging
from xml.dom import minidom

class Diccionario:

    def __init__(self):

        self.voc = {}

        self.LeeVoc(self.leeRegistro("LanguageID"))



    def LeeVoc(self,langID):

        # Esta función incorpora al diccionario global voc los contenidos
        # del archivo de idioma vocfile. voc es global para facilitar su
        # uso por todas las funciones.

        vocfile = "export.py.voc." + langID

        logging.info("Leyendo los archivos de idioma " + vocfile)

        xmldoc = minidom.parse(vocfile)

        words = xmldoc.getElementsByTagName('word')

        for word in words:
            
            try:
                self.voc[word.attributes['key'].value.encode('iso-8859-1')] = word.attributes['value'].value.encode('iso-8859-1')

            except:
                txt = "Error leyendo el voc"
                print txt
                logging.info(txt)


    def leeRegistro(self,variable):

        DEV = 0

        if DEV:

            valor = {'LanguageID': "1034", 'ExportPath': "C:\\Genomica\\Data\\CARExport\\"}[variable]

        else:

            try:

                keyHandle = win32api.RegOpenKeyEx(win32con.HKEY_LOCAL_MACHINE,"Software\\Genomica",0,win32con.KEY_ALL_ACCESS)

                valor,typeId = win32api.RegQueryValueEx(keyHandle,variable)

                win32api.RegCloseKey(keyHandle)

            except win32api.error, e:

                valor = {'LanguageID': "1034", 'ExportPath': "C:\\Genomica\\Data\\CARExport\\"}[variable]

        return valor


        

