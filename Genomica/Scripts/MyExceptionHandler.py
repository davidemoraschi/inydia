# -*- coding: cp1252 -*-
# Esta clase define la actuaci�n ante cualquier excepci�n
#
# El objeto recibe en su inicializaci�n una excepci�n.
# Mediante su m�todo act() se llama al m�todo que tenga por nombre
# handle + nombre_de_excepci�n, o a handleDefault en su defecto
# para hacer lo que corresponda



import traceback
import logging
import sys


class MyExceptionHandler:

    def __init__(self,e):

        self.e = e

    def act(self):

        nombreMetodo     = "handle" + self.e.__class__.__name__
        nombrePorDefecto = "handleDefault"

        if hasattr(self,nombreMetodo):

            metodo = getattr(self,nombreMetodo)

        else:

            metodo = getattr(self,nombrePorDefecto)
            
        metodo()


    def handleDefault(self):

        logging.error("**************************")
        logging.error(self.e.__class__.__name__)
        #logging.error(self.e.argE)
        logging.error("**************************")


        tmp = traceback.format_exception(sys.exc_type,sys.exc_value,sys.exc_traceback)
        logging.error("\n".join(tmp))

        
        
