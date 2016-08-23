# -*- coding: cp1252 -*-
# En este módulo defino las excepciones propias de la aplicación
# Por ahora, heredan sin más de la clase Exception


class MyException(Exception):

    def __init__(self):

        Exception.__init__(self)
        self.myErrorMessage = "internal.error"

    def getMyErrorMessage(self):

        return self.myErrorMessage
    



class ErrorAlEnviar(MyException):

    def __init__(self,e = None):

        MyException.__init__(self)
        self.argE = e
        self.myErrorMessage = "internal.error"


       
class ErrorAlSeleccionarCarpeta(MyException):

    def __init__(self,e = None):

        MyException.__init__(self)
        self.argE = e
        self.myErrorMessage = "internal.error"


        
class NoHayTestsEnLaCarpeta(MyException):

    def __init__(self,e = None):

        MyException.__init__(self)
        self.argE = e
        self.myErrorMessage = "internal.error"


        
class ErrorAlAbrirArchivoDeModulos(MyException):

    def __init__(self,e = None):

        MyException.__init__(self)
        self.argE = e
        self.myErrorMessage = "internal.error"


        
class ErrorAlLeerArchivoDeModulos(MyException):

    def __init__(self,e = None):

        MyException.__init__(self)
        self.argE = e
        self.myErrorMessage = "internal.error"


        
class ErrorAlLeerEscribirEnCola(MyException):

    def __init__(self,e = None):

        MyException.__init__(self)
        self.argE = e
        self.myErrorMessage = "internal.error"


        
class ErrorInputOrigen(MyException):

    def __init__(self,e = None):

        MyException.__init__(self)
        self.argE = e
        self.myErrorMessage = "internal.error"


        
class ErrorAlExportar(MyException):

    def __init__(self,e = None):

        MyException.__init__(self)
        self.argE = e
        self.myErrorMessage = "internal.error"


        
class ErrorAlExportarTests(MyException):

    def __init__(self,e = None):

        MyException.__init__(self)
        self.argE = e
        self.myErrorMessage = "internal.error"


        
class ErrorAlExportarTxt(MyException):

    def __init__(self,e = None):

        MyException.__init__(self)
        self.argE = e
        self.myErrorMessage = "internal.error"



class RutaDeExportacionNoEncontrada(MyException):

    def __init__(self,e = None):

        MyException.__init__(self)
        self.argE = e
        self.myErrorMessage = "exportpath.notfound"

        


