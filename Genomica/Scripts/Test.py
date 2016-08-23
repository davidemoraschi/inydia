# -*- coding: cp1252 -*-

import re
import os
import os.path
import sys
import time
from xml.dom import minidom
import zipfile



class Test:

    # La clase Test representa un test de un kit de Genomica, de cualquiera de sus plataformas
    # El objeto se inicializa con una trayectoria, tanto de Linux como de Windows, que debe ser siempre absoluta. Si no lo
    # es, lanza una expceción. Esta trayectoria debe llegar directamente hasta el test en el caso del ATS, o hasta las
    # carpetas Genomica y Sensovation en el caso del CAR. Esta trayectoria se convierte en el Id del test.
    # La clase permite realizar las operaciones que se describen a continuación
    #
    # Acceso:
    # - getPaths(self): Devuelve las trayectorias a los diferentes archivos que componen un test en un diccionario
    # - getInfoxml(self): Devuelve la cadena XML del info.xml si la ha leído o None en caso contrario
    # - getResultxml(self): Devuelve la cadena XML del result.xml si la ha leído o None en caso contrario
    # - getId(self): Devuelve la trayectoria hasta el test. Esta trayectoria es absoluta y llega hasta los archivos en el ATS
    #   y hasta la carpeta Genomica en el CAR
    #
    # Lectura:
    # - readInfoxml(self): Lee la cadena XML del info.xml
    # - readResultxml(self): Lee la cadena XML del info.xml
    # - readSVxml(self): Lee la cadena XML del xml de Sensovation
    #
    # Parseado
    # - parseInfoxml(self,campos): Devuelve diccionario con los contenidos de info.xml
    # - parseProbesResultxml(self): Devuelve diccionario con las sondas del result.xml
    # - parseVirusResultxml(self): Devuelve diccionario con los virus del result.xml
    # - parseControlsResultxml(self): Devuelve diccionario con las controles del result.xml
    # - parseSummaryResultxml(self): Devuelve cadena resumen del result.xml
    #
    # Operaciones
    # - createReports(self,assayid,lang): Genera informes a partir del result.xml de acuerdo al resultscript
    #   cuyo puntero recibe, y al idioma indicado. Solo funciona con el CAR??
    # - pack(self,target)
    # - reanalyse(self,assayModule,lang): Genera result.xml e informes a partir de acuerdo al resultscript
    #   cuyo puntero recibe, y al idioma indicado. Sólo funciona en el CAR, porque necesita el xml de Sensovation
    # - isAligned(self): Devuelve True o False según la imagen esté alineada


    def __init__(self,testPath):
        
        # utilizamos la ubicacion que me ha pasado como ID del test...
        self.id        = ""

	if testPath[0] <> '/' and testPath[1] <> ":":

		raise Exception("Las rutas deben ser absolutas!!!!!!")

        self.paths     = {    'resultxml'   : "",
                              'infoxml'     : "",
                              'bmpPath'     : "",
                              'reportPath'  : "",
                              'svxmlPath'   : "",
                              'iccal'       : ""
                         }

        self.infoxml   = ""
        self.resultxml = ""

	exists = 0

	try:

            self.id = testPath

            posiblePaths = { 'resultxml'    : {  'regex' : ["result.xml"],
                                                 'paths' : [testPath,testPath + os.sep + "Genomica"],   },
                             
                             'infoxml'      : {  'regex' : ["info.xml"],
                                                 'paths' : [testPath,testPath + os.sep + "Genomica"],   },

                             'iccal'        : {  'regex' : [".+\.iccal$"],
                                                 'paths' : [testPath,testPath + os.sep + "Genomica"],   },
                             
                             'bmpPath'      : {  'regex' : [".+\.bmp$"],
                                                 'paths' : [testPath,testPath + os.sep + "Genomica",testPath + os.sep + "Sensovation"],   },
                             
                             'reportPath'   : {  'regex' : ["result.prn.html"],
                                                 'paths' : [testPath,testPath + os.sep + "Genomica"],   },

                             'svxmlPath'    : {  'regex' : ["^[A-Z][0-9]{1,2}\.xml$"],
                                                 'paths' : [testPath + os.sep + "Sensovation",testPath + os.sep + "Clondiag",testPath + os.sep + "Genomica"],         },
                             }

            # para cada uno de los ficheros que queremos mapear
            for f in self.paths.keys():
                
                # para cada uno de sus posibles nombres
                for r in posiblePaths[f]['regex']:

                    regex = re.compile(r)
                    
                    # para cada una de sus posibles ubicaciones
                    for p in posiblePaths[f]['paths']:

                        # si existe
                        if os.path.isdir(p):

                            # miro dentro...
                            for i in os.listdir(p):

                                match = regex.search(i)

                                # y si lo encuentro lo guardo.
                                if match:
				    exists += 1
                                    self.paths[f] = p + os.sep + match.group(0)

	    if not exists:

		    raise Exception

	except:

		raise Exception("Test no encontrado")


    def getPaths(self):

        return self.paths


    def getInfoxml(self):

        return self.infoxml


    def getResultxml(self):

        return self.resultxml


    def getId(self):

        return self.id


    def readInfoxml(self):

        try:
            self.infoxml = minidom.parse(self.paths['infoxml'])

        except IOError, e:
            raise IOError("No encuentro el info.xml")


    def readSVxml(self):

        try:
            self.svxml = minidom.parse(self.paths['svxmlPath'])

        except IOError, e:
            raise IOError("No encuentro el xml de Sensovation")


    def readResultxml(self):

        try:

            self.resultxml = minidom.parse(self.paths['resultxml'])

        except IOError, e:
            raise IOError("No encuentro el result.xml")
    


    def parseInfoxml(self,campos = ["param_chipid",           "param_batchpos",     "param_testid",
                                    "param_processingstart",  "param_devicename",   "param_assayid",
                                    "param_testref",          "param_deviceserial", "param_position",
                                    "param_runid",            "param_carserial",    "param_date"]):

        # Devuelve un diccionario con los valores de los campos pasados como argumento

        if not self.infoxml:    self.readInfoxml()

        info = {}

        try:
            for i in campos:

                try:
                    info[i] = self.infoxml.firstChild.getElementsByTagName(i)[0].firstChild.data

                except:
                    info[i] = ''

            try:
                if not info["param_date"]:

                    info["param_date"] = time.ctime(float(info["param_processingstart"]))

            except:
                pass

            #try:
            #    regex = re.compile("[0-9]{4}_[0-9]{1,2}_[0-9]{1,2}_[0-9]{1,2}_[0-9]{1,2}_[0-9]{1,2}_([0-9]+)")
            #    match = regex.search(info["param_deviceserial"])
            #
            #    if match:
            #
            #        info["param_deviceserial"] = match.group(1)
            #
            #except:
            #    pass
            

        except Exception, e:
            raise Exception("Error al parsear info.xml")

        return info


    def parseProbesResultxml(self):

        # Devuelve un diccionario con los valores de las sondas leidas del result.xml

        if not self.resultxml:   self.readResultxml()

        sondas = {}

        try:
            sustancias = self.resultxml.getElementsByTagName('RESULTSET')

            for s in sustancias:

                if s.attributes['type'].value == 'substance':

                    sustancia = s.attributes['name'].value

                    values = s.getElementsByTagName('VALUE')

                    for v in values:

                        if v.attributes['name'].value == 'nVal':

                            sondas[sustancia] = v.attributes['value'].value

        except Exception, e:
            raise Exception("Error al parsear info.xml")


        return sondas


    def parseVirusResultxml(self):

        # Devuelve un diccionario con los resultados de los virus leidos del result.xml

        if not self.resultxml:   self.readResultxml()

        virus = {}

        try:
            sustancias = self.resultxml.getElementsByTagName('RESULTSET')

            for s in sustancias:

                if s.attributes['type'].value == 'virus':

                    sustancia = s.attributes['name'].value

                    virus[sustancia] = { 'value':   "",
                                         'control': "" }


                    for c in s.childNodes:

                        if c.nodeName == 'VALUE' and \
                           c.attributes['valueSubtype'].value == 'summarized substances':

                            virus[sustancia]['value'] = c.attributes['value'].value

                        if c.nodeName == 'VALUE' and \
                           c.attributes['valueSubtype'].value == 'associated controls':

                            virus[sustancia]['control'] = c.attributes['value'].value

        except Exception, e:
            raise Exception("Error al parsear info.xml")

        return virus


    def parseSummaryResultxml(self):

        # Devuelve el resumen leido del result.xml

        if not self.resultxml:   self.readResultxml()

        summary = ""

        try:
            frase = self.resultxml.getElementsByTagName('SUMMARY')

            if frase:
                
                summary = frase[0].attributes['value'].value

            if not summary:

                virus = self.parseVirusResultxml()
                ctrls = self.parseControlsResultxml()

                for v in virus:

                    if virus[v]['value'] == "positive" or virus[v]['value'] == "XXXPOSITIVEXXX":

                        summary += v + ", "

                    if virus[v]['value'] == "Uncertainty" or virus[v]['value'] == "XXXUNCERTAINTYXXX":

                        summary += "(" + v + "), "

                for c in ctrls:

                    summary += "[" + ctrls[c] + "]"

            if not summary:

                errValues = self.resultxml.getElementsByTagName('VALUE')

                for errValue in errValues:

                    if errValue.attributes['name'].value == 'errKey':

                        summary += "XXX" + errValue.attributes['value'].value + "XXX"

        except Exception, e:
            raise Exception("Error al parsear result.xml para summary")

        return summary
            

    def parseControlsResultxml(self):

        # Devuelve un diccionario con los resultados de los controles leidos del result.xml

        if not self.resultxml:   self.readResultxml()

        controls = {}

        try:
            sustancias = self.resultxml.getElementsByTagName('RESULTSET')

            for s in sustancias:

                if s.attributes['type'].value == 'control':

                    sustancia = s.attributes['name'].value

                    for c in s.childNodes:

                        if c.nodeName == 'VALUE' and \
                           c.attributes['valueType'].value == 'classification':

                            controls[sustancia] = c.attributes['value'].value

        except Exception, e:
            raise Exception("Error al parsear info.xml")

        return controls


    def createReports(self,assayModule,lang):

        tmp = os.getcwd().split(os.sep)
        address = os.sep.join(tmp[:-1]) + os.sep + "assays" + os.sep

	if type(assayModule) <> type(re):   # si no me ha pasado el modulo resultscript

            sys.path.append(address + assayModule)

	    try:
		    reload(assay)

	    except:
		    import resultscript as assay

	else:
		assay = assayModule
        
        assay.AssaysPath = address
        assay.idioma = lang

        tmp = self.paths['resultxml'].split(os.sep)
        OutputPath = os.sep.join(tmp[:-1]) + os.sep

        assay.Informes(OutputPath,
                       self.paths['bmpPath'],
                       self.paths['resultxml'],
                       self.paths['infoxml'])

    def pack(self,target = ''):

	# target indica el destino donde debe crearse el zip
	# - Si no se especifica nada, se crea en la trayectoria guarda en self.id (la accesible mediante self.getId()
	# - Si se especifica una carpeta, debe ser absoluta, y se crea ahi con el nombre deviceserial.testid.zip
	# - Si se especifica una ruta absoluta a un archivo, se crea ahi con ese nombre
	#
	# El metodo devuelve la localizacion, con ruta absoluta, del zip



	# Quiero que el zip contenga toda la estructura genomica/data/results/....
	# Por ello, me situo en results

	curDir = os.getcwd()
	
	separador = "genomica" + os.sep + "data" + os.sep + "results" + os.sep
	[before,after] = self.getId().split(separador)
	os.chdir(os.path.join(before,separador))

	# Si no ha puesto nada, creo el zip en self.id con el nombre deviceseria.testid.zip, y
	# si ha omitido solo el nombre del zip y ha puesto la carpeta le pongo ese nombre, y
	# si ha puesto un nombre a una ruta que no existe, lanzo Excepcion


	if not target:

		dirname = self.getId()
		info = self.parseInfoxml()
		zipname = info['param_deviceserial'] + "." + info['param_testid'] + ".zip"


	elif os.path.isdir(target):
		
		dirname = target
		info = self.parseInfoxml()
		zipname = info['param_deviceserial'] + "." + info['param_testid'] + ".zip"

	else:
		dirname,zipname = os.path.split(target)
		if not os.path.isdir(dirname):	raise Exception


	# Estoy dentro de genomica/data/results. Creo el zip

	#print "Estoy en " + os.getcwd() + " y creo " + zipname

	myZip = zipfile.ZipFile(zipname,"w")
	for root,dirs,files in os.walk(after):

		for f in files:

			if f <> zipname:

				#print "Escribiendo " + root + os.sep + f

				myZip.write(os.path.join(root,f))

	myZip.close()

	# y lo muevo donde deberia estar

	os.rename(os.path.join(before,separador,zipname),os.path.join(dirname,zipname))

	# y me vuelvo a donde estaba

	os.chdir(curDir)

	#print "Estoy en Test.py y he creado el zip " + os.path.join(dirname,zipname)

	return os.path.join(dirname,zipname)


    def zipForGenomicaLims(self,comprimido):
        # Los archivos zip para subir al LIMS de Genomica deben contener todos los archivos en una
        # única carpeta que debe llamarse Genomica si es de un CAR, y de cualquier otro modo si es
        # un ATS. Por ejemplo, los contenidos de un test de un CAR serían:
        #
        # Genomica\A1.bmp
        # Genomica\A1.xml
        # Genomica\result.raw.html
        # ...
        #
        
        # Creo el archivo zip
	myZip = zipfile.ZipFile(comprimido,"w")


	# Me cambio al lugar donde debo estar
	cwd = os.getcwd()
	os.chdir(self.getId())


	# Determino el nombre de la carpeta que irá dentro del zip
	folderInsideZip = ""
	if self.parseInfoxml()['param_devicename'] == "CAR":

            folderInsideZip = "Genomica"

        else:

            folderInsideZip = self.parseInfoxml()['param_deviceserial'] + "." + self.parseInfoxml()['param_testid']



        # Determino los archivos que voy a meter y cómo los voy a llamar
        filesInZip = []

	if self.parseInfoxml()['param_devicename'] == "CAR":

            try:
                filesInZip = [(os.path.join("Genomica",f),os.path.join(folderInsideZip,f)) for f in os.listdir("Genomica")]
                filesInZip += [(os.path.join("Sensovation",f),os.path.join(folderInsideZip,f)) for f in os.listdir("Sensovation")]

            except:
                pass

        if not filesInZip:

            filesInZip = [(f,os.path.join(folderInsideZip,f)) for f in os.listdir(".")]



        # Meto los archivos
        for (fileName,fileInZip) in filesInZip:

            try:

                myZip.write(fileName.encode('iso-8859-1'),fileInZip.encode('iso-8859-1'))

            except Exception,e:

                logging.info("Error al escribir en el zip el archivo " + fileName)

        # Cierro el zip
        myZip.close()


    def reanalyse(self,assayModule,lang):

	if not (os.path.exists(self.paths['infoxml']) and os.path.exists(self.paths['svxmlPath'])):

		raise Exception("No encuentro la imagen o/y los datos")
        
        tmp = os.getcwd().split(os.sep)
        address = os.sep.join(tmp[:-1]) + os.sep + "assays" + os.sep

	if type(assayModule) <> type(re):

                assayID = assayModule

		sys.path.append(address + assayModule)
		try:
			reload(assay)
		except:
			import resultscript as assay
	else:
		assay = assayModule
		assayID = assayModule.___file__.split(os.sep)


	self.changeAssayIdInInfoxml(assayID)

	assay.AssaysPath = address
	assay.idioma = lang

	tmp = self.paths['infoxml'].split(os.sep)
	OutputPath = os.sep.join(tmp[:-1]) + os.sep

        print OutputPath
	return assay.main(
				OutputPath,
				lang,
				'',
				self.paths['bmpPath'],
				self.paths['svxmlPath'],
				'')


    def changeAssayIdInInfoxml(self,assayID):
        # Esta función cambia el assayID del archivo info.xml al especificado como argumento
        # Se utiliza antes de reanalizar porque se guiará por lo que diga el info.xml para
        # cargar el resultscript correspondiente

        # Si no tengo el info.xml lo consigo
        info = self.parseInfoxml()

        # Si hace falta
        if assayID <> info["param_assayid"]:

            # Lo modifico
            self.infoxml.getElementsByTagName("param_chipid")[0].firstChild.data = "0000000" + assayID
            self.infoxml.getElementsByTagName("param_assayid")[0].firstChild.data = assayID

            # Y lo guardo para la próxima
            fh = file(self.paths['infoxml'],"w")
            fh.write(self.infoxml.toxml())
            fh.close()


    def isAligned(self):

        info = self.parseInfoxml()
        
        if info['param_devicename'] == 'ATS':

            return True

        else:

            self.readSVxml()
            status = self.svxml.getElementsByTagName('Status')
            ErrorID = int(status[0].attributes['ErrorID'].value)
            
            return ErrorID == 0
            

        

        

