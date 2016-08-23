# -*- coding: cp1252 -*-


from threading import *
from MyExceptions import *
from MyExceptionHandler import *
from xml.dom import minidom
import time
import re
import os
import shutil
import os.path
import win32api
import win32con
import win32gui
import logging



class ExportThread(Thread):

    def __init__(self,tests,config):

        Thread.__init__(self)
        
        self.tests    = tests
        self.config   = config
        self.abort    = False
        self.progress = 0


    def notifyToWindow(self,notifyWindow):

        self.notifyWindow = notifyWindow


    def getProgress(self):

        return self.progress


    def run(self):

        '''
        # Stub para pruebas, muy útil.

        try:

            for i in range(10):
                time.sleep(0.5)
                if self.abort: break       # La tarea debe poder abortarse así
                print "*",
                self.progress = float(i+1)/10   # El progreso se guarda aquí
                self.notifyWindow.notifyProgress(self.progress)  # y se notifica así
                #if i > 3: a = tmp[5]  # Genera errores artificial

        except Exception,e:

            self.progress = -1                             # El progreso se guarda en -1
            self.notifyWindow.notifyProgress(self.progress)  # Notifica el error a la ventana
            eh = MyExceptionHandler(ErrorAlExportar(e))    # Gestiona la excepción
            eh.act()            

           
        '''

        count           = len(self.tests)
        n               = 1

        dictTxt            = self.config.lee("txt")
        dictTests          = self.config.lee("tests")

        exportTxt          = dictTxt["status"] == "on"
        exportTests        = dictTests["status"] == "on"
        exportTxtPath      = dictTxt["exportPath"]
        exportTestsPath    = dictTests["exportPath"]

        registryExportPath = self.getExportPath()


        # Entonces aquí tenemos el exportPath definido en tres sitios
        # - exportTxtPath es la ubicación destino para los export.txt especificado en la configuración de export.py
        # - exportTestsPath es la ubicación destino para las carpetas de los tests especificado en la configuración de export.py
        # - registryExportPath es la ubicación especificada en el registro para la varible ExportPath

        # En este caso, pasamos de la configuración del export.py y cogemos la del registro para todo. Si no existe
        # daremos error

        exportTxtPath = registryExportPath
        exportTestsPath = registryExportPath

        # Fin del problema


        if exportTxtPath[-1] <> "\\":               exportTxtPath += "\\"
        if exportTestsPath[-1] <> "\\":             exportTestsPath += "\\"

        linkRegex = re.compile("\.\./Sensovation/")


        runreports = []
        


        for t in self.tests:

            info = t.parseInfoxml()

            expTestException = None
            expTxtException  = None

            if exportTests:

                try:

                    tPaths   = t.getPaths()
                    infoPath = tPaths['infoxml']
                    source   = "\\".join(infoPath.split("\\")[:-1])
                    target   = exportTestsPath + "\\" + info['param_deviceserial'] + "." + \
                               info['param_position'] + "." + info['param_testref']
                    image    = tPaths['bmpPath'].split("\\")[-1]

                    try:
                        runreport = [os.path.join(source.split(info['param_deviceserial'])[0],
                                                  info['param_deviceserial']),
                                     info['param_deviceserial']]

                        if not runreport in runreports:     runreports.append(runreport)

                    except:
                        logging.error("Error al preparar la copia del runreport")
                    


                    if not os.path.exists(target):

                        shutil.copytree(source,target)

                        if info['param_devicename'] == 'CAR':

                            shutil.copy(tPaths['bmpPath'],target + "\\" + image)

                        lee = file(target + "\\result.prn.html","r")
                        texto = lee.read()
                        lee.close()

                        texto = linkRegex.sub("",texto)

                        escribe = file(target + "\\result.prn.html","w")
                        escribe.write(texto)
                        escribe.close()

                except Exception,e:

                    if not os.path.exists(target):

                        expTestException = RutaDeExportacionNoEncontrada(e)

                    else:

                        expTestException = ErrorAlExportarTests(e)

                    #expTestException = e
                    #raise ErrorAlExportarTests(e)

            if exportTxt:

                try:

                    txtFileName = exportTxtPath + "export"
                    addTxt      = ""

                    if info['param_devicename'] == 'CAR':

                        data = info['param_deviceserial']
                        dateTxt = ""
                        runidTxt = ""

                        if dictTxt['includeDateInName'] == "on":
                            
                            fecha = ".".join(data.split("_")[:3])
                            hora = ".".join(data.split("_")[3:6])
                            dateTxt = " " + fecha + " " + hora

                        if dictTxt['includeRunIdInName'] == "on":

                            runidTxt = " " + data.split("_")[-1]

                        addTxt = dateTxt + runidTxt

                    txtFileName += addTxt + ".txt"

                    self.guardaCadenaResultados(self.traduce(info['param_assayid'],
                                                             self.produceCadenaResultados(t,dictTxt)),
                                                txtFileName)

                except Exception,e:

                    if not os.path.exists(exportTxtPath):

                        expTxtException = RutaDeExportacionNoEncontrada(e)

                    else:

                        expTxtException = ErrorAlExportarTests(e)


            if self.abort: break


            self.progress = float(n)/count   # El progreso se guarda aquí
            self.notifyWindow.notifyProgress(self.progress)  # y se notifica así
            n = n + 1

            if expTestException:

                self.progress = expTestException                
                self.notifyWindow.notifyProgress(self.progress)  # Notifica el error a la ventana

                # Bug: esto nunca se ejecuta????
                eh = MyExceptionHandler(ErrorAlExportarTests(e))    # Gestiona la excepción
                eh.act()
                expTestException = None

                
            if expTxtException:

                self.progress = expTxtException                 # El progreso se guarda en -1
                self.notifyWindow.notifyProgress(self.progress)  # Notifica el error a la ventana

                # Bug: esto nunca se ejecuta????
                eh = MyExceptionHandler(ErrorAlExportarTxt(e))    # Gestiona la excepción
                eh.act()
                expTxtException = None


        if runreports:

            for r in runreports:

                try:
                    shutil.copy(os.path.join(r[0],"runreport.html"),os.path.join(exportTestsPath,r[1] + ".runreport.html"))
                    shutil.copy(os.path.join(r[0],"logo_cabecera1.gif"),os.path.join(exportTestsPath,"logo_cabecera1.gif"))

                except:
                    logging.error("Error al copiar el runreport")
                


    def kill(self):

        self.abort = True


    def leeRegistro(self,variable):

        DEV = 0

        if DEV:

            valor = {'LanguageID': "1034", 'ExportPath': "C:\\Genomica\\Data\\CARExport\\"}[variable]

        else:

            keyHandle = win32api.RegOpenKeyEx(win32con.HKEY_LOCAL_MACHINE,"Software\\Genomica",0,win32con.KEY_ALL_ACCESS)

            valor,typeId = win32api.RegQueryValueEx(keyHandle,variable)

            win32api.RegCloseKey(keyHandle)

        return valor


    def getExportPath(self):

        try:
            exportPath = self.leeRegistro("ExportPath")

        except:

            raise RutaDeExportacionNoEncontrada

            '''if os.path.exists("C:\\Genomica\\Data\\CARExport\\"):

                exportPath = "C:\\Genomica\\Data\\CARExport\\"

            else:

                try:
                    os.mkdir("C:\\Genomica\\Data\\CARExport")
                    exportPath = "C:\\Genomica\\Data\\CARExport\\"

                except:
                    exportPath = "C:\\Genomica\\Data\\"'''

        return exportPath


    def produceCadenaResultados(self,test,config):

        metadata = test.parseInfoxml()
        data     = test.parseVirusResultxml()
        ctrl     = test.parseControlsResultxml()
        resumen  = test.parseSummaryResultxml()
        

        sep      = config['sep']
        eol      = re.compile('\n').sub(config['eol'],"\n")
        

        salida = ""
        faltante = "-100"     # este será el valor de dato faltante


        # Datos de la muestra (10 campos)

        datos = ["param_assayid","param_chipid","param_position",
                  "param_processingstart","param_runid","param_testref","param_carserial","param_deviceserial"]

        metadata["param_processingstart"] = time.strftime("%d/%m/%Y",time.localtime(float(metadata["param_processingstart"])))

        for d in datos:

            try:
                salida += metadata[d] + sep

            except:
                salida += faltante + sep

        relleno = ["HHH","III","JJJ","KKK","LLL","MMM","NNN","OOO","PPP","QQQ","RRR","SSS"]
                

        for i in xrange(11):

            if i > len(datos):     salida += relleno[i] + sep


        
        # Resultados (50 campos)

        code = {   "XXXPOSITIVEXXX":            "1",
                   "XXXNEGATIVEXXX":            "0",
                   "XXXUNCERTAINTYXXX":         "-1",
                   "positive":                  "1",
                   "negative":                  "0",
                   "Uncertainty":               "-1"
                   }


        virus = [i for i in data.keys() if i <> 'XXXignoreXXX']
        virus.sort()

        if not virus: virus = [faltante]*50

        for v in virus:

            try:
                salida += code[data[v]['value']] + sep

            except:
                salida += faltante        + sep


        for i in xrange(51):

            if i > len(virus):     salida += "" + sep


        # Controles (10 campos)

        code = {   "XXXPASSEDXXX":              "0",
                   "XXXFAILEDXXX":              "1",
                   "XXXNOSIGNALXXX":            "2",
                   "XXXUNCERTAINTYXXX":         "-1",
                   "Uncertainty":               "-1",
                   "No concluyente":            "-1",
                   "Conforme":                  "0",
                   "Passed":                    "0",
                   "Sin señal":                 "2",
                   "No signal":                 "2",
                   "No conforme":               "1",
                   "Failed":                    "1",
                   }

        controles = ctrl.keys()
        controles.sort()

        if not controles: controles = [faltante]*10

        for c in controles:

            try:
                salida += code[ctrl[c].encode('latin-1')] + sep

            except:
                salida += faltante + sep

        for i in xrange(11):

            if i > len(controles):     salida += "" + sep


        # Observaciones (10 campos)

        if resumen.find("XXXINVALIDXXX") >=0:

            valido = "0"
            diagnostico = "-1"
    
        else:
            valido = "1"

            if resumen.find("XXXNEGATIVEXXX") >=0:

                diagnostico = "0"

            else:

                diagnostico = "1"

        salida += resumen     + sep
        salida += valido      + sep
        salida += diagnostico + sep
        salida += "AAA"          + sep
        salida += "BBB"          + sep
        salida += "CCC"          + sep
        salida += "DDD"          + sep
        salida += "EEE"          + sep
        salida += "FFF"          + sep
        salida += "GGG"          + eol
        
        return salida


    def guardaCadenaResultados(self,cadena,archivo):

        cabeceraCadena = ",".join(cadena.split(",")[:10])
        nuevoTxt = []

        if os.path.exists(archivo):

            fh = file(archivo,"r")
            todo = unicode(fh.read(),'latin-1')
            fh.close()
            lineas = [i for i in todo.split("\n") if len(i) > 0]

            cambio = False

            for linea in lineas:

                cabeceraLinea = ",".join(linea.split(",")[:10])

                if cabeceraCadena <> cabeceraLinea:

                    nuevoTxt.append(linea)
                
                else:

                    cambio = True
                    nuevoTxt.append(cadena)

            if not cambio:

                    nuevoTxt.append(cadena)

        else:

            nuevoTxt.append(cadena)

            
        fh = file(archivo,"w")
        fh.write("\n".join(nuevoTxt).encode('latin-1'))
        fh.close()


    def traduce(self,ensayo,texto):

        try:
            xmldoc = minidom.parse(
                os.path.join(os.path.splitdrive(self.leeRegistro("InstallDir"))[0] + os.sep,
                             "Genomica","Data","Assays",ensayo,
                             ensayo + ".voc." + self.leeRegistro("LanguageID")))

        except IOError,e:
            return texto
        

        words = xmldoc.getElementsByTagName('word')
        voc = {}

        for word in words:
            
            try:
                voc[word.attributes['key'].value] = word.attributes['value'].value

            except:
                pass

  	for k in voc.keys():

            tmp = ''
            while tmp <> texto:

                tmp = texto
                texto = re.sub ( "XXX" + k + "XXX", voc[k], texto, re.MULTILINE | re.DOTALL)

        return texto

