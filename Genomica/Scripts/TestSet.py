# -*- coding: latin-1 -*-


import Test
import Translator

import logging
import re
import os
import os.path
import zipfile
import time
from sets import Set



class TestSet:

    #    En el constructor se recibe como argumento una ruta o una lista de rutas donde
    # se deben buscar tests. Se escanearán todas las carpetas y subcarpetas en su interior
    # hasta encontrarlos
    # Posteriormente, se pueden añadir más de la misma forma con el método addTests,
    # o quitarlos con removeTests.
    #
    # Al iniciar el objeto con una o más rutas o añadir más, se llama al método findTests.
    # Este método busca los tests y pone sus trayectorias absolutas en la variable self.testPaths.
    # Estas trayectorias son las que llegan hasta los archivos del test en el sistema ATS
    # y hasta las carpetas Genomica y Sensovation en el sistema CAR. Pueden obtenerse con
    # el método getTestPaths()
    #
    # Métodos:
    # - addTests
    # - removeTests
    # - getTestPaths
    # - findTests
    #
    #    Con las trayectorias a los tests en testPaths, el método buildSet() construye una lista de objetos 
    # de la clase Test en self.tests. La cuenta de cuántos tests tenemos se obtiene con getSetCount().
    # La lista de objetos Test se obtiene con getTests(). El método buildSetLastOnly() incluye sólo el
    # último reanálisis de cada test.
    #
    # Métodos:
    # - buildSet
    # - buildSetLastOnly
    # - buildSetLastAlignedOnly
    # - getSetCount
    # - getTests
    #
    #   Con un set construido podemos crear informes de datos o resultados, formatearlos al gusto u
    # obtenerlos en formato excel.
    #
    # - createRawView: Crea el tag xml <Worksheet> correspondiente a la hoja de datos crudos con todos
    #   los tests del testset. Devuelve una lista de cadenas, que son las líneas de texto.
    #   Esto se hace así porque es más eficiente que manejar cadenas de texto largas.
    # - createResultView: Crea el tag xml <Worksheet> correspondiente a la hoja de resultados con todos
    #   los tests del testset. Devuelve una lista de cadenas, que son las líneas de texto.
    #   Esto se hace así porque es más eficiente que manejar cadenas de texto largas.
    # - formatView: Recibe como argumento una cadena que es un xml con elementos <Worksheet> y le pone
    #   la cabecera y el pie para convertirse en un xml de Excel. Es el encargado de definir los estilos.
    #   Devuelve la cadena de texto xml lista para guardar.
    # - getFullView: Devuelve cadena xml del Excel con la hoja de datos crudos y de resultados
    # - getResultView: Devuelve cadena xml del Excel con la hoja de resultados
    # - getRawView: Devuelve cadena xml del Excel con la hoja de datos crudos
    # - pack: Crea un zip con los zips generados de cada test producidos por su propio método pack
    #   Ver comentarios en el código sobre los argumentos que recibe
    # - createReports: Llama uno a uno para todos los tests del set al método createReports
    # - reanalyse: Llama uno a uno para todos los tests del set al método reanalyse
    #
    # Algunos de estos métodos se sirven de otros auxiliares como
    # - transKeys


    filesInRealPath   = ['info.xml','result.xml','result.prn.html','result.raw.html','result.res.html']
    

    def __init__(self,arguments):

        self.tests        = []
        self.testPaths    = []
        self.myTranslators = {}

        self.addTests(arguments)


    def addTests(self,arguments):
        
        # Este metodo busca en la direccion o direcciones pasadas como argumento
        # los tests que contegan, donde quiera que esten, y mete sus paths
        # en la variable miembro de tipo lista testPaths

        if type(arguments) == type(["lista"]):

            candidatos = arguments

        else:

            candidatos = [arguments]

        searchLocations = [d for d in candidatos if os.path.isdir(d)]

	for s in searchLocations:

	    if s[0] <> '/' and s[1] <> ':':

		    raise Exception("Las rutas deben ser absolutas!!!!!!")
	
	logging.debug("Considero los rutas " + ", ".join(searchLocations))

        self.testPaths.extend(self.findTests(searchLocations))


    def removeTests(self,arguments):

        # Este metodo busca en la direccion o direcciones pasadas como argumento
        # los tests que contegan, donde quiera que esten, y quita sus paths
        # de la variable miembro de tipo lista testPaths, caso de que esten ahi.

        if type(arguments) == type(["lista"]):

            candidatos = arguments

        else:

            candidatos = [arguments]

        searchLocations = [d for d in candidatos if os.path.isdir(d)]

        for t in self.findTests(searchLocations):

            try:
                self.testPaths.remove(t)

            except:
                pass
            

    def getTestPaths(self):

        # Este metodo devuelve una lista con los paths a los tests que
        # le han metido al objeto mediante los metodos addTests y removeTests

	logging.debug("Tengo los tests " + ", ".join(self.testPaths))



        return self.testPaths
    

    def findTests(self,searchLocations):

        # Este metodo devuelve los paths a los tests contenidos en la direccion
        # o direcciones pasadas en la lista searchLocations
	# Se asume que las rutas pasadas son absolutas o relativas desde la carpeta
	# donde esta este módulo.. Es preferible que sea todo absoluto, asi no hay problemas.
        #
        # - En el caso del ATS es la ruta a la carpeta que contiene los archivos
        # listados en la variable de clase filesInRealPath
        # - En el caso del CAR es la ruta a la carpeta que contiene las carpetas
        # Genomica y Sensovation

        locations = []

        for loc in searchLocations:

            for root,dirs,files in os.walk(loc):

                for f in self.filesInRealPath:

                    if f in files:

                        locations.append(root)
                        break



        if locations:
            
            regex = re.compile("Genomica$")

            for i in xrange(len(locations)):

                locations[i] = regex.sub("",locations[i])
                if locations[i][-1] == os.sep:  locations[i] = locations[i][:-1]

        return locations


    def getSetCount(self):

        # Devuelve el numero de tests que contiene el Set
        # (no el numero de paths añadidos)

        return len(self.tests)


    def buildTestSet(self,inclusionCriteria):

        # Primero creo una lista de objetos Test con todos los paths que tengo
        # para luego ir quitando lo que no quiero.
        # Esto no parece eficiente pero sí lo es, si piensas que de
        # todos tengo que abrir el info.xml!!!
        self.allTests = []
        
        for p in self.testPaths:

            try:
                self.allTests.append(Test.Test(p))

            except:
                logging.info("Error al intentar construir el objeto test para la ruta " + p)



        # Hago un diccionario donde las claves son los identificadores únicos de
        # test del tipo "deviceserial.testid" y los valores son una lista de todos
        # los tests del set que corresponden a esa identificación
        uniqueTests = {}

        for t in self.allTests:

            try:
                
                info = t.parseInfoxml()

                try:    uniqueTests[info['param_deviceserial'] + "." + info['param_testid']].append(t)
                except: uniqueTests[info['param_deviceserial'] + "." + info['param_testid']] = [t]

            except:
                pass


        # Ahora tengo que hacer el filtrado
        # Los valores por defecto para los criterios de inclusión son:
        criteria = inclusionCriteria.keys()
        if not 'includeNotAligned' in criteria:     inclusionCriteria['includeNotAligned'] = 'yes'
        if not 'includeLastOnly' in criteria:       inclusionCriteria['includeLastOnly'] = 'yes'



        # Hago el filtrado

        # Para cada identificador único
        for k in uniqueTests.keys():


            # Criterio includeLastOnly
            if inclusionCriteria['includeLastOnly'] == 'yes':

                # Hago una lista de los reanálisis que inicializo con 0 para que
                # el máx(reanalysisList) no dé error nunca
                reanalysisList = [0]

                
                # Voy a ir pasando de uno en uno por todos y
                # guardando el candidato aquí
                lastTest  = []

                for t in uniqueTests[k]:

                    testId = t.getId()
                    try:
                        reanalysis = int(testId.split(os.sep)[-1])

                    except:
                        reanalysis = 1

                    if reanalysis > max(reanalysisList): lastTest = [t]
                    reanalysisList.append(reanalysis)

                # Ya lo tengo, lo guardo sobreescribiendo la lista de todos los que había antes
                uniqueTests[k] = lastTest


            # Criterio includeNotAligned
            if inclusionCriteria['includeNotAligned'] == 'no':

                # Si están alineados los incluyo y si no, no
                uniqueTests[k] = [t for t in uniqueTests[k] if t.isAligned()]


        # y finalmente uno todos los tests en self.tests
        for k in uniqueTests.keys():

            self.tests.extend(uniqueTests[k])



    def buildSet(self):

        inclusionCriteria = { 'includeNotAligned':      'yes',
                              'includeLastOnly':        'no'
                              }

        self.buildTestSet(inclusionCriteria)

        
    def buildSetLastOnly(self):
        
        inclusionCriteria = { 'includeNotAligned':      'yes',
                              'includeLastOnly':        'yes'
                              }

        self.buildTestSet(inclusionCriteria)

        
    def buildSetLastAlignedOnly(self):

        inclusionCriteria = { 'includeNotAligned':      'no',
                              'includeLastOnly':        'yes'
                              }

        self.buildTestSet(inclusionCriteria)


    def getTests(self):

	return self.tests


    def transKeys(self,dictionary,myTranslator):

        # Este metodo traduce las claves de un diccionario
        # utilizando el objeto de la clase Translator pasado
        # como argumento.

        oldkeys = dictionary.keys()
        newkeys = myTranslator.translate(oldkeys)

        dict_tmp = {}

        for i in xrange(len(oldkeys)):

            dict_tmp[newkeys[i]] = dictionary[oldkeys[i]]

        return dict_tmp
        

    def createRawView(self,lang = '1033'):

        # Crea el tag Worksheet correspondiente a los datos crudos
        # del Excel xml

	logging.debug("Inicio la crecion del RawView")

        tabla      = {}
        errores    = {}

        info_keys  = Set([])
        probe_keys = Set([])

        # Este for crea el diccionario tabla. Las claves son los Ids de los tests y los valores
        # los valores son otro diccionario. En este diccionario, las claves son las cabeceras de
        # las columnas de lo que luego será la hoja del excel con los datos crudos.
        # Estas cabeceras incluyen las claves del infoxml y los nombres de las sondas.
        #
        # Las claves también salen en sendas listas, info_keys y probe_keys. Es importante
        # destacar que estas listas contienen todas las claves vistas en todos los tests, aunque
        # haya tests que para esas claves no tengan valor. Son listas extensivas.
        #
        # Los tests que generan un error en este paso van al diccionario errores. Las claves vuelven
        # a ser los Ids de los tests y los valores deberían ser los errores que han dado!!!!!
        

        for test in self.tests:

            try:

                Id    = test.getId()
                info  = test.parseInfoxml(["param_chipid","param_batchpos","param_testid","param_date",
			"param_devicename","param_assayid","param_testref","param_deviceserial"])
                probe = test.parseProbesResultxml()

		logging.debug("Voy a hacer " + Id)

                try:
                    myTranslator = self.myTranslators[info['param_assayid'][-5:]]

                except:
                    self.myTranslators[info['param_assayid'][-5:]] = Translator.Translator("vocs" + os.sep + info['param_assayid'][-5:] + ".voc." + lang)
                    myTranslator = self.myTranslators[info['param_assayid'][-5:]]

                info  = self.transKeys(info,myTranslator)
                probe = self.transKeys(probe,myTranslator)

                info_keys  = info_keys | Set(info.keys())
                probe_keys = probe_keys | Set(probe.keys())

                tabla[Id] = {}

                for i in info_keys:   tabla[Id][i] = info[i]
                for p in probe_keys:  tabla[Id][p] = probe[p]

		logging.debug("Terminado")

            except IOError,e:

                logging.error("Error 1")
                errores[test.getId()] = e

            except Exception,e:

                logging.error("Error 2: " + test.getId())
                errores[test.getId()] = e


        # Ordeno, y así las llamo para todos en el mismo orden
        info_keys = [i for i in info_keys]
        info_keys.sort()

        probe_keys = [i for i in probe_keys]
        probe_keys.sort()


        output = []

        output.append("<Worksheet ss:Name=\"Datos crudos\">\n")

        '''info_keys = [i for i in info_keys]
        info_keys.sort()

        probe_keys = [i for i in probe_keys]
        probe_keys.sort()

        output = ""

        output += "<Worksheet ss:Name=\"Datos crudos\">\n"'''

        # Ahora voy a hacer la cabecera de la tabla de excel. Voy metiendo primero info_keys y luego probe_keys en sendos for
        output.append("<Table x:FullColumns=\"1\" x:FullRows=\"1\">\n")

        cabecera = "<Row>\n"
        anchuras = ""
        for i in xrange(len(info_keys)):

            if info_keys[i] == myTranslator.translate("param_chipid"):

                anchuras += "<Column ss:Index=\"" + str(i+1) + "\" ss:AutoFitWidth=\"0\" ss:Width=\"100\"/>\n"

            if info_keys[i] == myTranslator.translate("param_processingstart"):
                
                anchuras += "<Column ss:Index=\"" + str(i+1) + "\" ss:AutoFitWidth=\"0\" ss:Width=\"140\"/>\n"
               
            cabecera += "<Cell ss:StyleID=\"s26\"><Data ss:Type=\"String\">" + info_keys[i] + "</Data></Cell>\n"
            
        for p in probe_keys:    cabecera += "<Cell ss:StyleID=\"s23\"><Data ss:Type=\"String\">" + p + "</Data></Cell>\n"
        cabecera += "</Row>\n"

        output.append(anchuras + cabecera)

	logging.debug("Ya tengo la cabecera")
                                

        # Ahora voy escribiendo las filas de los tests
        for t in tabla.keys():

            fila = []

            fila.append("<Row>\n")

            for i in info_keys:

                try:
                    fila.append("<Cell ss:StyleID=\"s24\"><Data ss:Type=\"String\">" + tabla[t][i] + "</Data></Cell>\n")

                except:
                    fila.append("<Cell ss:StyleID=\"s27\"><Data ss:Type=\"String\">Error!!</Data></Cell>\n")
                    

            for p in probe_keys:

                try:
                    fila.append("<Cell ss:StyleID=\"s22\"><Data ss:Type=\"Number\">" + tabla[t][p] + "</Data></Cell>\n")

                except:
                    fila.append("<Cell ss:StyleID=\"s27\"><Data ss:Type=\"String\">Error</Data></Cell>\n")

            fila.append("</Row>\n")

            output.extend(fila)

	logging.debug("Ya tengo los tests")


        # Ahora voy a por los errores
        for e in errores.keys():

            output.append("<Row>\n")
            output.append("<Cell ss:StyleID=\"s28\"><Data ss:Type=\"String\">" + e + "</Data></Cell>\n")
            output.append("<Cell ss:StyleID=\"s28\"><Data ss:Type=\"String\">" + e + "</Data></Cell>\n")
            output.append("</Row>\n")

	logging.debug("y los errores")
                

        output.extend(["</Table>\n",
                  "<WorksheetOptions xmlns=\"urn:schemas-microsoft-com:office:excel\">",
                  "<PageSetup>\n",
                  "<Header x:Margin=\"0\"/>\n",
                  "<Footer x:Margin=\"0\"/>\n",
                  "<PageMargins x:Bottom=\"0.984251969\" x:Left=\"0.78740157499999996\"\n",
                  "x:Right=\"0.78740157499999996\" x:Top=\"0.984251969\"/>\n",
                  "</PageSetup>\n",
                  "<Print>\n",
                  "<ValidPrinterInfo/>\n",
                  "<PaperSizeIndex>9</PaperSizeIndex>\n",
                  "<HorizontalResolution>600</HorizontalResolution>\n",
                  "<VerticalResolution>600</VerticalResolution>\n",
                  "</Print>\n",
                  "<Selected/>\n",
                  "<Panes>\n",
                  "<Pane>\n",
                  "<Number>3</Number>\n",
                  "<ActiveRow>1</ActiveRow>\n",
                  "</Pane>\n",
                  "</Panes>\n",
                  "<ProtectObjects>False</ProtectObjects>\n",
                  "<ProtectScenarios>False</ProtectScenarios>\n",
                  "</WorksheetOptions>\n",
                  "</Worksheet>\n"])

	logging.debug("Termino el RawView")
                    
        return output
        

    def createResultView(self):

        # Crea el tag Worksheet correspondiente a los resultados
        # del Excel xml
        # Ver los comentarios en el método anterior porque es idéntico a este

	logging.debug("Inicio la creacion del ResultView")

        tabla      = {}
        errores    = {}

        info_keys  = Set([])


        for test in self.tests:

	    try:

                Id      = test.getId()
                info  = test.parseInfoxml(["param_chipid","param_batchpos","param_testid","param_date",
			"param_devicename","param_assayid","param_testref","param_deviceserial"])
                summary = test.parseSummaryResultxml()

                try:
                    myTranslator = self.myTranslators[info['param_assayid'][-5:]]

                except:
                    self.myTranslators[info['param_assayid'][-5:]] = Translator.Translator("vocs" + os.sep + info['param_assayid'][-5:] + ".voc.1033")
                    myTranslator = self.Translators[info['param_assayid'][-5:]]

                info    = self.transKeys(info,myTranslator)

                info_keys  = info_keys | Set(info.keys())

                tabla[Id] = {}

                for i in info.keys():   tabla[Id][i] = info[i]

                tabla[Id]['summary'] = myTranslator.translate(summary)

            except IOError,e:

                errores[test.getId()] = e.message

            except Exception,e:

		    errores[test.getId()] = e.message

        info_keys = [i for i in info_keys]
        info_keys.sort()

        output = []

        output.append("<Worksheet ss:Name=\"Resultados\">\n")
        output.append("<Table x:FullColumns=\"1\" x:FullRows=\"1\">\n")

        cabecera = "<Row>\n"
        anchuras = ""
        for i in xrange(len(info_keys)):

            if info_keys[i] == myTranslator.translate("param_chipid"):

                anchuras += "<Column ss:Index=\"" + str(i+1) + "\" ss:AutoFitWidth=\"0\" ss:Width=\"100\"/>\n"

            if info_keys[i] == myTranslator.translate("param_processingstart"):
                
                anchuras += "<Column ss:Index=\"" + str(i+1) + "\" ss:AutoFitWidth=\"0\" ss:Width=\"140\"/>\n"
                
            cabecera += "<Cell ss:StyleID=\"s26\"><Data ss:Type=\"String\">" + info_keys[i] + "</Data></Cell>\n"
            
        cabecera += "<Cell ss:StyleID=\"s23\"><Data ss:Type=\"String\">Result</Data></Cell>\n"
        cabecera += "</Row>\n"

        output.append(anchuras + cabecera)


        for t in tabla.keys():

            fila = []

            fila.append("<Row>\n")

            for i in info_keys:

                try:
                    fila.append("<Cell ss:StyleID=\"s24\"><Data ss:Type=\"String\">" + tabla[t][i] + "</Data></Cell>\n")

                except:
                    fila.append("<Cell ss:StyleID=\"s27\"><Data ss:Type=\"String\">Error</Data></Cell>\n")
                    

            try:
                fila.append("<Cell ss:StyleID=\"s25\"><Data ss:Type=\"String\">" + tabla[t]['summary'] + "</Data></Cell>\n")

            except Exception,e:
           
                fila.append("<Cell ss:StyleID=\"s27\"><Data ss:Type=\"String\">Error</Data></Cell>\n")

            fila.append("</Row>\n")

            output.extend(fila)


        for e in errores.keys():

            output.append("<Row>\n")
            output.append("<Cell ss:StyleID=\"s28\"><Data ss:Type=\"String\">" + "Error" + "</Data></Cell>\n")
            output.append("</Row>\n")
                

        output.extend(["</Table>\n",
                  "<WorksheetOptions xmlns=\"urn:schemas-microsoft-com:office:excel\">",
                  "<PageSetup>\n",
                  "<Header x:Margin=\"0\"/>\n",
                  "<Footer x:Margin=\"0\"/>\n",
                  "<PageMargins x:Bottom=\"0.984251969\" x:Left=\"0.78740157499999996\"\n",
                  "x:Right=\"0.78740157499999996\" x:Top=\"0.984251969\"/>\n",
                  "</PageSetup>\n",
                  "<Print>\n",
                  "<ValidPrinterInfo/>\n",
                  "<PaperSizeIndex>9</PaperSizeIndex>\n",
                  "<HorizontalResolution>600</HorizontalResolution>\n",
                  "<VerticalResolution>600</VerticalResolution>\n",
                  "</Print>\n",
                  "<Selected/>\n",
                  "<Panes>\n",
                  "<Pane>\n",
                  "<Number>3</Number>\n",
                  "<ActiveRow>1</ActiveRow>\n",
                  "</Pane>\n",
                  "</Panes>\n",
                  "<ProtectObjects>False</ProtectObjects>\n",
                  "<ProtectScenarios>False</ProtectScenarios>\n",
                  "</WorksheetOptions>\n",
                  "</Worksheet>\n"])

	logging.debug("Termino el ResultView")
                    
        return output


    def formatView(self,texto):

        # Crea Excel xml con las hojas que recibe como argumento en texto

	logging.debug("Empiezo a darle el formato")

        cabecera =  ["<?xml version=\"1.0\" encoding=\"iso-8859-1\"?>\n",
                    "<?mso-application progid=\"Excel.Sheet\"?>\n",
                    "<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"\n",
                    " xmlns:o=\"urn:schemas-microsoft-com:office:office\"\n",
                    " xmlns:x=\"urn:schemas-microsoft-com:office:excel\"\n",
                    " xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\"\n",
                    " xmlns:html=\"http://www.w3.org/TR/REC-html40\">\n",
                    " <DocumentProperties xmlns=\"urn:schemas-microsoft-com:office:office\">\n",
                    "  <Author>Microsoft Corporation</Author>\n",
                    "  <LastAuthor>jproman</LastAuthor>\n",
                    "  <Created>1996-11-27T10:00:04Z</Created>\n",
                    "  <LastSaved>2003-04-16T10:36:01Z</LastSaved>\n",
                    "  <Version>11.5606</Version>\n",
                    " </DocumentProperties>\n",
                    " <ExcelWorkbook xmlns=\"urn:schemas-microsoft-com:office:excel\">\n",
                    "  <WindowHeight>4500</WindowHeight>\n",
                    "  <WindowWidth>9420</WindowWidth>\n",
                    "  <WindowTopX>120</WindowTopX>\n",
                    "  <WindowTopY>132</WindowTopY>\n",
                    "  <AcceptLabelsInFormulas/>\n",
                    "  <ProtectStructure>False</ProtectStructure>\n",
                    "  <ProtectWindows>False</ProtectWindows>\n",
                    "  <DisplayInkNotes>False</DisplayInkNotes>\n",
                    " </ExcelWorkbook>\n"]

        # Estilos
        # s22 => Numeros
        # s23 => Encabezados de datos
        # s24 => Textos centrado
        # s25 => Textos alineados a la izquierda
        # s26 => Encabezados de info.xml
        # s27 => Celda con error
        # s28 => Mensaje de error

        cabecera.extend(["<Styles>\n",
                    "  <Style ss:ID=\"Default\" ss:Name=\"Normal\">\n",
                    "   <Borders/>\n",
                    "   <Font/>\n",
                    "   <Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Bottom\"/>\n",
                    "  </Style>\n",
                    "  <Style ss:ID=\"s22\">\n",
                    "   <Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Bottom\"/>\n",
                    "  </Style>\n",
                    "  <Style ss:ID=\"s23\">\n",
                    "   <Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Bottom\"/>\n",
                    "  </Style>\n",
                    "  <Style ss:ID=\"s24\">\n",
                    "   <Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Bottom\"/>\n",
                    "  </Style>\n",
                    "  <Style ss:ID=\"s25\">\n",
                    "   <Alignment ss:Horizontal=\"Left\" ss:Vertical=\"Bottom\"/>\n",
                    "  </Style>\n",
                    "  <Style ss:ID=\"s26\">\n",
                    "   <Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Bottom\"/>\n",
                    "   <Borders>\n",
                    "    <Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"2\"/>\n",
                    "   </Borders>\n",
                    "   <Font x:Family=\"Swiss\" ss:Bold=\"1\"/>\n",
                    "  </Style>\n",
                    "  <Style ss:ID=\"s27\">\n",
                    "   <Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Bottom\"/>\n",
                    "  </Style>\n",
                    "  <Style ss:ID=\"s28\">\n",
                    "   <Alignment ss:Horizontal=\"Left\" ss:Vertical=\"Bottom\"/>\n",
                    "  </Style>\n",
                    " </Styles>\n"])
        
        pie =      ["</Workbook>\n"]

	logging.debug("Termino el formateo")

        return "".join(cabecera + texto + pie)



    def getFullView(self):

	# Crea el Excel xml con las dos hojas (datos crudos y resultados)

	logging.debug("Ha pedido vista completa")

        return self.formatView(self.createRawView() + self.createResultView())

    
    def getRawView(self):

        # Crea el Excel xml sólo con la hoja de datos crudos

	logging.debug("Ha pedido solo datos")

        return self.formatView(self.createRawView())


    def getResultView(self):

        # Crea el Excel xml con sólo la hoja de resultados

	logging.debug("Ha pedido solo resultados")

        return self.formatView(self.createResultView())


    def pack(self,target='',mode = 'w'):

	# Crea un zip con todos los tests del set

	# El argumento target es el destino del zip que queremos crear
	# - Si no ha especificado nada, lo creo en el directorio actual, y por nombre pongo la fecha
	# - Si ha especificado una carpeta, lo guardo en ella y por nombre pongo la fecha
	# - Si ha especificado un nombre con trayectoria absoluta lo guardo ahi.

	if not target:

		dirname = os.getcwd()
		zipname = time.strftime("%d %b %Y %H.%M.%S") + ".zip"


	elif os.path.isdir(target):

		dirname = target
		zipname = time.strftime("%d %b %Y %H.%M.%S") + ".zip"

	else:
		dirname,zipname = os.path.split(target)
		if not os.path.isdir(dirname): raise Exception



	# Primeamente voy test por test llamando al metodo pack. Le paso la carpeta
	# destino, dirname, para que cree los zips individuales ahi mismo
	# En zipnames me quedo con una lista de todas las rutas absolutas a estos zips.


	zipnames = []

	for i in self.tests:

		zipnames.append(i.pack(dirname))


	curDir = os.getcwd()
	os.chdir(dirname)

	#print "Me muevo a " + dirname + " para crear " + zipname

	try:
		myZip = zipfile.ZipFile(zipname,mode)

	except:
		myZip = zipfile.ZipFile(zipname,'w')

	
	for i in [os.path.split(j)[1] for j in zipnames]:

		#print "Anyado " + i

		myZip.write(i.encode('utf-8'))

	myZip.close()

	for i in zipnames:

		#print "Borro " + i

		os.unlink(i)

	os.chdir(curDir)

	#print "He creado " + os.path.join(dirname,zipname)

	return os.path.join(dirname,zipname)


    def createReports(self,assayModule,lang):

	for i in self.tests:

		i.createReports(assayModule,lang)

    def reanalyse(self,assayModule,lang):

	output = 1    

	for i in self.tests:

		try:
                    output *= i.reanalyse(assayModule,lang)

                except:
                    print "Error al reanalizar " + i.getId()
		    logging.debug("Error al reanalizar " + i.getId())

	return output



	
