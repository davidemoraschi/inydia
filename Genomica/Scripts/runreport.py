# -*- encoding: latin-1 -*-


from xml.dom import minidom
import string
import time
import re
import logging
import shutil
import os
import os.path
import win32api
import win32con


import traceback       
import sys

voc = {}
WorkingListAssayPath = ''




############################################################
#
#  Funciones de servicio
#
############################################################





def Guarda(destino,texto):

        # Guarda la cadena texto en el archivo destino, con la codificación latin-1

        logging.info("Guardando " + str(destino))

        output = file(destino,"w")
        output.write(texto.encode('latin-1'))
        output.close





def Lee(filename):

        # Lee el archivo filename y lo devuelve como una única cadena de texto

        logging.info("Leyendo " + str(filename))

	input = file(filename,"r")
	return "".join(input.readlines())



###################################################################################
#
#   Eleccion del reanalisis
#
###################################################################################



def SortStringsAsNumbers(x,y):

    if    int(x)>int(y): return 1
    elif  int(x)<int(y): return -1
    else:            return 0





def UltimoReanalisis(TestPath):

        # Devuelve el path al último reanálisis de un test
        #
        # Por ejemplo, para C:\Genomica\Data\Results\2008_1_4_12_12_34_456\C3,
        # si ha sido reanalizada una sola vez, devuelve,
        # C:\Genomica\Data\Results\2008_1_4_12_12_34_456\C3\2
        

        lista = [i for i in os.listdir(TestPath) if os.path.isdir(TestPath + "\\" + i)]

        lista.sort(SortStringsAsNumbers)

        return TestPath + "\\" + lista[-1]



###################################################################################
#
#   Lectura de datos en XML
#
###################################################################################



def GetXMLElement(xmldoc,element):
        
        # element = [['RESULTSET','type','virus'],['RESULTSET','type','substance'],['VALUE','','']]
        
        if len(element)>0:

                ByTags0 = xmldoc.getElementsByTagName(element[0][0])

                elements0 = []
                
                for i in ByTags0:

                    if element[0][1]:

                        if i.attributes[element[0][1]].value == element[0][2]:  elements0.append(i)

                    else:  elements0 = ByTags0

                if len(element)>1:

                    elements1 = []

                    for e0 in elements0:

                        hijos0 = []

                        for hijo in e0.childNodes:

                            try:
                                    if hijo.tagName == element[1][0]: hijos0.append(hijo)
                                    
                            except:
                                    pass

                             
                        for j in hijos0:

                            if element[1][1]:

                                if j.attributes[element[1][1]].value == element[1][2]:   elements1.append(j)

                            else:  elements1.append(j)

                    if len(element)>2:

                        elements2 = []

                        for e1 in elements1:

                            hijos1 = []

                            for hijo in e1.childNodes:

                                    try:
                                            if hijo.tagName == element[2][0]: hijos1.append(hijo)

                                    except:
                                            pass
                                            

                            for k in hijos1:

                                if element[2][1]:

                                    if k.attributes[element[2][1]].value == element[2][2]:   elements2.append(k)

                                else:  elements2.append(k)

                        return elements2

                    else:  return elements1
                else:      return elements0
        else:              return 0





def GetResultDOM(xml) :

	from xml.dom import minidom
	

	resultdoc = minidom.parseString(xml)
	result = []
	for child in resultdoc.documentElement.childNodes:
		if child.nodeName == "RESULTS":
			result.append(child)
	return result





def ParseSensovationInfoXML(infoxml):

        # Parsea el info.xml y devuelve diccionario
    

        logging.info("Parseando info.xml")
    
        rawxmldoc = minidom.parseString(infoxml)

        SVmetadata = {}

        campos = ["param_chipid","param_processingstart","param_wellid","param_carserial","param_deviceserial",
                  "param_runid","param_position","param_assayid","param_testref","param_imagename"]


        for i in campos:

            try:
                SVmetadata[i] = rawxmldoc.firstChild.getElementsByTagName(i)[0].firstChild.data

            except:
                SVmetadata[i] = ''

        try:
                
            tmp = float(SVmetadata['param_processingstart'])
            import time
            SVmetadata['param_processingstart'] = time.strftime("%d/%m/%Y",time.localtime(float(SVmetadata['param_processingstart'])))

        except:
           pass

        if not SVmetadata['param_assayid'] or not SVmetadata['param_testref'] or not SVmetadata['param_position']:

                msg = 'No encontrados los datos necesarios dentro del info.xml'
                logging.error(msg)
                raise Exception(msg)
         
        return SVmetadata





def ParseResultXML(resultxml):

        # Este parser extrae del result.xml de Clondiag-Sensovation la información de los datos
        # crudos y la frase resumen. Devuelve una lista donde el primer elemento es un diccionario
        # con los primeros, y el segundo es una cadena con lo segundo.
        #
        # Se utiliza sólo para extraer los datos


        logging.info("Leyendo result.xml")


        results = GetResultDOM(resultxml)

        results = results[0]

        data = {}
        ctrl = {}


        #substances       = GetXMLElement(results,[['RESULTSET','type','virus'],['RESULTSET','type','substance']])
        virus            = GetXMLElement(results,[['RESULTSET','type','virus']])
        controls         = GetXMLElement(results,[['RESULTSET','type','control'],['RESULTSET','type','substance']])


        for s in virus:

            values = s.childNodes

            for v in values:

                if v.nodeName == 'VALUE' and \
                   v.attributes['name'].value == 'qVal' and \
                   v.attributes['valueSubtype'].value == 'summarized substances':

                    data[s.attributes['name'].value] = v.attributes['value'].value


        for c in controls:

            values = c.childNodes

            for va in values:

                if va.nodeName == 'VALUE' and \
                   va.attributes['name'].value == 'qVal':

                    ctrl[c.attributes['name'].value] = va.attributes['value'].value


        summarytag = results.getElementsByTagName('SUMMARY')

        return [data,ctrl,summarytag[0].attributes['value'].value]



def ParseResultXMLyInfoXML(infoxml,resultxml):

    # Esta función es un refrito de ParseResultXML e InfoXML, que devuelve
    # una terna con los metadatos, los datos crudos y el resumen
    # Los dos primeros son diccionarios y el último es una cadena
    #
    # Me viene bien así

    metadata = ParseSensovationInfoXML(infoxml)
                
    [data,ctrl,summary] = ParseResultXML(resultxml)

    return [metadata,data,ctrl,summary]








#############################################################
#
#    voc
#
#############################################################





def LeeVoc(vocfile):

        # Esta función incorpora al diccionario global voc los contenidos
        # del archivo de idioma vocfile. voc es global para facilitar su
        # uso por todas las funciones.

        global voc

        logging.info("Leyendo los archivos de idioma " + vocfile)

        xmldoc = minidom.parse(vocfile)

        words = xmldoc.getElementsByTagName('word')

        for word in words:
            
            try:
                voc[word.attributes['key'].value] = word.attributes['value'].value

            except:
                txt = "Error leyendo el voc"
                print txt
                logging.info(txt)





def Voz(key,dict):

        # Esta función facilita el uso de diccionarios en varias funciones
        # Si key existe pasa su valor y si no lo devuelve tal cual
        # Por defecto, el diccionario es el global voc 

	try:
		valor = dict[key]

	except:
		valor = key

	return valor





def Traduce(txt):

        # Esta función traduce la cadena txt a un idioma con los
        # términos definidos en el diccionario global voc
        # Utiliza mi función para reemplazar


        return Remplaza(txt)




#################################################################################
#
#  Mi Remplace
#
#################################################################################





def Remplaza(texto,dict = {},flag = ''):

        # En texto reemplaza las palabras clave por los valores definidos en dict
        # Si no se adjunta dict, se entiende que es voc y que el flag es XXX
        # Si se adjunta dict y no flag, se entiende que el flag es ''
        # Utiliza la función Voz para evitar errores.

        if not dict:

            global voc
            dict = voc
            flag = 'XXX'


  	for k in dict.keys():

                tmp = ''
                while tmp <> texto:

                        tmp = texto
                        texto = re.sub ( flag + k + flag, Voz(k,dict), texto, re.MULTILINE | re.DOTALL)

        return texto



#################################################################################
#
#    Exportación al LIMS de Genomica
#
#################################################################################


def EnviaAlLims(ensayo):

        try:
                import LimsUploadCar
                LimsUploadCar.EnviaDesdeLector(ensayo)

        except Exception,e:

                msg = "\n" + "\n".join(traceback.format_exception(sys.exc_type,sys.exc_value,sys.exc_traceback))
                    
                logging.error("Error al buscar el LIMS de Genomica!!! " + msg)
                




#################################################################################
#
#    Exportación a un LIMS ajeno
#
#################################################################################


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



def GetExportPath():

        exportPath = ""

        try:
                exportPath = LeeRegistro()

                logging.info("Se leyó del registro de Windows el ExportPath")

                if not os.path.exists(exportPath):

                        logging.info("El exportPath no existe!!!!")
                        raise IndexError    #    :-p

        except:
                logging.info("No se pudo leer ExportPath del registro o no existe!!")

                candidateExportPath = os.path.join(
                        os.path.splitdrive(LeeRegistro("InstallDir"))[0] + os.sep,
                        "genomica","data","CARExport") + os.sep

                #candidateExportPath = "C:\\genomica\\data\\CARExport\\"

                if os.path.exists(candidateExportPath):

                        exportPath = candidateExportPath

                else:
                        logging.info("No está CARExport pero intento crearlo")
                        
                        try:
                                os.mkdir(candidateExportPath)
                                exportPath = candidateExportPath

                                logging("He creado " + candidateExportPath)

                        except:
                                loggin.info("No he podido crear " + candidateExportPath)
                                
                                exportPath = "C:\\Genomica\\Data\\"

        #exportPath += "export.txt"

        logging.info("ExporPath = " + exportPath)

        return exportPath


                                


def ProduceCadenaResultados(metadata,data,ctrl,resumen):

        salida = ""

        sep      = ","        # caracter de separacion
        faltante = "-100"     # este será el valor de dato faltante
        eol      = "\n"


        # Datos de la muestra (10 campos)

        datos = ["param_assayid","param_chipid","param_position",
                  "param_processingstart","param_runid","param_testref","param_carserial","param_deviceserial"]

        for d in datos:

                try:
                        salida += metadata[d] + sep

                except:
                        salida += faltante + sep

        for i in xrange(11):

                if i > len(datos):     salida += "" + sep


        
        # Resultados (50 campos)

        code = {   "XXXPOSITIVEXXX":            "1",
                   "XXXNEGATIVEXXX":            "0",
                   "XXXUNCERTAINTYXXX":         "-1"
                   }

        virus = [i for i in data.keys() if i <> 'XXXignoreXXX']
        virus.sort()

        if not virus: virus = [faltante]*50

        for v in virus:
                
                try:
                        salida += code[data[v]] + sep

                except:
                        salida += faltante        + sep

        for i in xrange(51):

                if i > len(virus):     salida += "" + sep


        # Controles (10 campos)

        code = {   "XXXPASSEDXXX":              "0",
                   "XXXFAILEDXXX":              "1",
                   "XXXNOSIGNALXXX":            "2",
                   "XXXUNCERTAINTYXXX":         "-1"
                   }

        controles = ctrl.keys()
        controles.sort()

        if not controles: controles = [faltante]*10

        for c in controles:

                try:
                        salida += code[ctrl[c]] + sep

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
        salida += ""          + sep
        salida += ""          + sep
        salida += ""          + sep
        salida += ""          + sep
        salida += ""          + sep
        salida += ""          + sep
        salida += ""          + eol

        year,month,day,hour,minute,sec,runnumber = metadata["param_deviceserial"].split("_")

        tag = year + "-" + month + "-" + day + " " + hour + "h " + minute + "m " + sec + "s " + runnumber
        
        return salida,"export " + tag + ".txt"



#################################################################################
#
#    Extracción
#
#################################################################################



def ProduceRunReportDict(metadata,resumen):

    return {"<!-- " + metadata['param_assayid'] + metadata['param_position'] + " -->": resumen}




    
###################################################################################
#
#   Sensovation's main
#
###################################################################################




def main(OutputPath, lang, RunID, DataString):

    # Esta función produce el report del run a partir del workinglist y
    # los datos de las muestras que recibe como argumento
    #
    # La función devuelve 1 si ha ido bien o 0 si ha habido problemas
    #
 

    try:

        logging.basicConfig(	level=logging.DEBUG,
	      	            	format='%(asctime)s %(levelname)s %(message)s',
				filename=OutputPath + "runreport.log.txt",
				filemode='a')

        logging.info("Empezamos main...")



        # Arreglo todos los argumentos...
      
        global WorkingListAssayPath
        global ExportPath
        global idioma

        WorkingListAssayPath = os.path.join(
                os.path.splitdrive(LeeRegistro("InstallDir"))[0] + os.sep,
                "Genomica","Data","scripts") + os.sep
        #WorkingListAssayPath = "C:\\genomica\\data\\scripts\\"
        #ExportPath           = "C:\\Genomica\\Data\\CARExport\\SinValidar\\"
        ExportPath           = GetExportPath()
        idioma               = str(lang)
        RunID                = str(RunID)
        OutputPath = OutputPath.replace(".\\","\\")
        OutputPath = OutputPath.replace("\\\\","\\")



        # Reporto al log lo que he recibido

        logging.info("OutputPath " + str(OutputPath))
        logging.info("lang " + idioma)
        logging.info("RunID " + RunID)
        logging.info("Data " + str(DataString))



        # Defino el resto de las variables que me hacen falta...        

        wlerror     = WorkingListAssayPath + "runreport.error.html"
        wltemplate  = OutputPath + "workinglist.html"
        runreport   = OutputPath + "runreport.html"
        vocfile     = WorkingListAssayPath + "workinglist.voc." + idioma
        logfile     = OutputPath + "runreport.log.txt"



        # Paso el DataString a samples...

        LeeVoc(vocfile)

        samples = []
        data = DataString.split(",")
        plateid = data[0]
        number  = data[1]

        j = 2
        for i in xrange(int(number)):

                if int(plateid):
                        samples.append([data[j],data[j+1],data[j+2],data[j+3]])
                        j += 4
                else:
                        samples.append([data[j],data[j+1],data[j+2]])
                        j += 3



        # Pregunto antes de nada por el LIMS de Genomica...

        ENVIALIMS = 0

        try:
                import LimsUploadCar
                import win32gui

                if 1==win32gui.MessageBox(0,"¿Quieres enviar ya los resultados al LIMS?\n" + \
                                          "Si prefieres reanalizar imágenes pulsa en Cancelar",\
                                          "Genomica LIMS",1):

                        ENVIALIMS = 1

        except Exception,e:
                
                msg = "\n" + "\n".join(traceback.format_exception(sys.exc_type,sys.exc_value,sys.exc_traceback))
                    
                logging.error("No encuentro el LIMS de Genomica!!!\n" + msg)



        # Voy una por una...

        ReportDict = {}
        ExportStrg = ''

        for sample in samples:

                # ... rellenando el runreport y exportando...

                position = sample[-3]
                assayid  = sample[-2][-5:]

                try:

                        LeeVoc(os.path.join(
                                os.path.splitdrive(LeeRegistro("InstallDir"))[0] + os.sep,
                                "Genomica","Data","Assays",assayid,assayid + ".voc." + idioma))
                        #LeeVoc("C:\\genomica\\data\\assays\\" + assayid + "\\" + assayid + ".voc." + idioma)

                        ultimoReanalisis = UltimoReanalisis(OutputPath + position)

                        infoxml   = Lee(ultimoReanalisis + "\\Genomica\\info.xml")
                        resultxml = Lee(ultimoReanalisis + "\\Genomica\\result.xml")

                        [metadata,data,ctrl,summary] = ParseResultXMLyInfoXML(infoxml,resultxml)

                        tmp = ProduceRunReportDict(metadata,summary)
                        key = tmp.keys()[0]
                        logging.info("Reportando " + key + " ==> " + tmp[key])
                        ReportDict[key] = tmp[key]

                        tmp2,exportFileName = ProduceCadenaResultados(metadata,data,ctrl,summary)
                        logging.info("Exportando " + tmp2)
                        ExportStrg += tmp2


                except Exception, e2:

                        msg = "\n" + "\n".join(traceback.format_exception(sys.exc_type,sys.exc_value,sys.exc_traceback))
                       
                        logging.error("Error al reportar " + position + msg)



                # ... y exportando al LIMS de Genomica si procede.

                if ENVIALIMS:

                        print UltimoReanalisis(OutputPath + position) + "\\Genomica"
                        EnviaAlLims(UltimoReanalisis(OutputPath + position) + "\\Genomica")
                                
                else:
                        logging.info("No había que enviar al LIMS de Genomica")



        # Ahora en ReportDict tengo las frases para el report
        # y en ExportStrg tengo el txt exportable al LIMS del cliente.
        # Guardo runreport, export y se acabó

        try:
                Guarda(os.path.join(ExportPath,exportFileName),Traduce(ExportStrg))
                ReportDict["<!-- EXPORTREPORT -->"] = "XXXEXPORTOKXXX"
                logging.info("Exportación realizada con éxito")
                
        except:
                ReportDict["<!-- EXPORTREPORT -->"] = "XXXEXPORTFAILUREXXX"
                logging.info("La exportación falló al guardar export.txt")


        reporthtml = Traduce(Remplaza(Lee(wltemplate),ReportDict))
        Guarda(runreport,reporthtml)                

        logging.info("Fin del runreport main. Exit = 1")

        return 1


    except Exception, e2:

        msg = "\n" + "\n".join(traceback.format_exception(sys.exc_type,sys.exc_value,sys.exc_traceback))

        msg = "Fin del runreport main. Exit = 0. " + msg

        logging.info(msg)

        return 0
	





if __name__ == "__main__":

        logging.basicConfig(	level=logging.DEBUG,
	       	            	format='%(asctime)s %(levelname)s %(message)s',
				filename="log.txt",
				filemode='a')

        logging.info("Empezamos ejecución...")


        def UltimaCarpeta():

                carpetas = os.listdir("c:\\Genomica\\Data\\Results")

                carpetas.sort(SortFolder)

                return carpetas[-1]


        def SortFolder(x,y):

                campos_x = x.split("_")
                campos_y = y.split("_")

                num_x = int(campos_x[0])*100000000 + int(campos_x[1])*1000000 + int(campos_x[2])*10000 + int(campos_x[3])*100 + int(campos_x[4])
                num_y = int(campos_y[0])*100000000 + int(campos_y[1])*1000000 + int(campos_y[2])*10000 + int(campos_y[3])*100 + int(campos_y[4])        
                
                if    num_x > num_y: return 1
                elif  num_x < num_y: return -1
                else:            return 0





        ultima       = UltimaCarpeta()
        #ultima       = "2008_2_20_13_14_17_976"   # si funciona
        #ltima       = "2008_1_4_20_41_8_36"

        OutputPath   = "C:\\Genomica\\Data\\Results\\" + ultima + "\\"
        RunID        = "Mi Run ID"
        lang         = "1033"
        num_samples = len([i for i in os.listdir(OutputPath[:-1]) if os.path.isdir(OutputPath + i)])
        tiras       = num_samples/8
        ATs         = 1
        ASs         = 0
        assayid     = '40204'

        print os.listdir(OutputPath)
        pocillos = [j for j in os.listdir(OutputPath) if os.path.isdir(OutputPath + j)]
        print pocillos

        samples  = ''
        
        letra = {1: 'A', 2: 'B', 3: 'C', 4: 'D', 5:'E', 6: 'F', 7: 'G', 8: 'H'}
        j=1
        k=1
        
        for i in xrange(num_samples):

                samples += ","
                
                if ASs:  samples += str(j) + ","
                j=j+8
                if j > 96:
                        k=k+1
                        j=k

                samples += pocillos[i] + ","
                samples += "123450" + str(i+1) + assayid + ","
                samples += pocillos[i]


        samples = str(num_samples) + samples

        if ATs: samples = "0," + samples
        else:   samples = "1," + samples

        print "Enviando al LIMS " + ultima + "\n\n"
        print samples
    
        print main(OutputPath,lang,RunID,samples)
        



