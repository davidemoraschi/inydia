# -*- coding: cp1252 -*-




from TestSet import *
from MyExceptionHandler import *
from MyExceptions import *
from ClassLoader import *

import logging
import os
import os.path
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


import sys
sys.path.append(os.path.join(
        os.path.splitdrive(LeeRegistro("InstallDir"))[0] + os.sep,
        "Genomica","Data","Scripts"))


def main(inputOrigen = "",modulesFile = "",mockArgument1 = "",mockArgument2 = ""):

    raiz = os.path.join(
        os.path.splitdrive(LeeRegistro("InstallDir"))[0] + os.sep,
        "Genomica","Data","Scripts")
    cwd = os.getcwd()
    os.chdir(raiz)



    logging.basicConfig(    level=logging.DEBUG,
                            format='%(asctime)s %(levelname)s %(message)s',
                            filename=os.path.join(raiz,"export.py.log.txt"),
                            filemode='a')

    logging.info("inputOrigen recibido = " + inputOrigen)
    logging.info("modulesFile recibido = " + modulesFile)
    logging.info("mockArgument1 recibido = " + mockArgument1)
    logging.info("mockArgument2 recibido = " + mockArgument2)

    #######################################################################################
    # Dónde estoy??

    # 1.- Si he recibido cuatro argumentos...
    if mockArgument2:

        # ... es que estoy como workinglist.py. Lo debo estar probando,
        # y por ello tengo que pasar de los argumentos que he recibido
        # y generar un inputOrigen para hacer las pruebas:
        inputOrigen = GeneraArgDev()

    else:

        # ... si no, soy export.py, y me han pasado dos. El inputOrigen
        # y el modulesFile son los buenos. No necesito hacer nada
        pass


    # 2.- En cualquiera de los casos compruebo si el segundo argumento es un
    # modulesFile válido y si no cojo el que está por defecto

    if not os.path.exists(modulesFile):     modulesFile = "appModules.xml"


    # 3.- Fin. Al final queda lo siguiente:

    logging.info("inputOrigen = " + inputOrigen)
    logging.info("modulesFile = " + modulesFile)
    
    #######################################################################################


    try:
        
        logging.info("Cargo origen")

        
        
        
        miLoader    = ClassLoader(modulesFile)

        

    
        claseOrigen = miLoader.load("origen")
        miOrigen    = claseOrigen()

        logging.info("Proceso origen")


        miOrigen.start(inputOrigen)
        carpeta = miOrigen.dameCarpeta()

        logging.info("Las carpetas son...")
        logging.info(carpeta)

      

        if carpeta:

            logging.info("Cargo buscador")

            claseBuscador = miLoader.load("buscador")
            miBuscador    = claseBuscador()

            logging.info("Proceso buscador")

            miBuscador.start(carpeta)
            inputSelector = miBuscador.dameTests()


            if inputSelector:

                logging.info("Cargo selector")

                claseSelector = miLoader.load("selector")
                miSelector    = claseSelector()
                logging.info("Proceso selector")

                miSelector.start(inputSelector)
                seleccionados =  miSelector.dameSeleccionados()

                logging.debug("Seleccionados = (debajo)")
                logging.debug(seleccionados)

                

                if seleccionados:

                    logging.info("Cargo exportador")

                    claseExportador = miLoader.load("exportador")

                    miExportador    = claseExportador()
       
                    inputExportador = seleccionados

                    logging.info("Proceso exportador")

                    miExportador.start(inputExportador)

                    miExportador.exporta()

        logging.info("Fin")

        os.chdir(cwd)

        return 1


    except Exception, e:

        eh = MyExceptionHandler(e)
        eh.act()

        os.chdir(cwd)

        return 0





def GeneraArgDev():

    runname1 = "2008_4_13_17_32_36_1215"

    raiz = "C:\\Genomica\\Data\\Results\\"

    pocillos1 = [i for i in os.listdir(raiz + runname1) \
                if os.path.isdir(raiz + runname1 + "\\" + i)]

    argumento = str(len(pocillos1))

    for i in pocillos1:

        argumento += "," + runname1 + "," + i

    return argumento
    



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

