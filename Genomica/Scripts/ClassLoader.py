# -*- coding: cp1252 -*-


from xml.dom import minidom
from MyExceptions import *
import logging


class ClassLoader():


    def __init__(self,modulesFile):

        try:
            self.configdoc = minidom.parse(modulesFile)

        except Exception,e:
            raise ErrorAlAbrirArchivoDeModulos(e)


    def load(self,tagname):

        try:
            logging.debug("Empiezo")
            nombreModulo       = self.configdoc.getElementsByTagName(tagname)[0].attributes["modulo"].value
            logging.debug("Debo cargar " + nombreModulo)
            modulo             = __import__(nombreModulo)
            logging.debug("Lo tengo aquí")
            logging.debug(modulo)
            clase              = getattr(modulo,nombreModulo)
            logging.debug("Ya tengo la clase")

            return clase

        except Exception,e:
            raise ErrorAlLeerArchivoDeModulos(e)


        

            
