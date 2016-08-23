# -*- encoding: latin-1 -*-


from xml.dom import minidom
import string
import time
import re
import logging
import shutil
import os.path
import os
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



def LeeRegistro(variable):

        try:
                import win32api
                import win32con

                keyHandle = win32api.RegOpenKeyEx(win32con.HKEY_LOCAL_MACHINE,"Software\\Genomica",0,win32con.KEY_ALL_ACCESS)

                try:
                        value,typeId = win32api.RegQueryValueEx(keyHandle,variable)

                except:
                        value,typeId = win32api.RegQueryValueEx(keyHandle,variable)

                win32api.RegCloseKey(keyHandle)

        except:
                value = ""

        return value





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

        # Reemplaza las palabras clave que contenga texto que estén definidas en dict por los valores
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



def SortByPosition(x,y):

    px = x[-3]
    py = y[-3]

    dict = {'A': 1, 'B': 2, 'C': 3, 'D':4, 'E': 5, 'F': 6, 'G': 7, 'H': 8}

    regex = re.compile("^([A-H]{1})([0-9]{1,2})$")

    matchx = regex.search(px)
    if matchx:
        letrax    = matchx.group(1)
        numerox   = matchx.group(2)
        posicionx = dict[letrax]+(int(numerox)-1)*8

        
    matchy = regex.search(py)
    if matchy:
        letray    = matchy.group(1)
        numeroy   = matchy.group(2)
        posiciony = dict[letray]+(int(numeroy)-1)*8


    return posicionx - posiciony
    




def CreaFilasMuestras(samples):
    
    samples.sort(SortByPosition)


    header = '''<table style="text-align: left; width: 100%;" border="1" cellspacing="0"
 cellpadding="3">
  <tbody>
    <tr>
      <td style="text-align: center; vertical-align: middle; width: 5%;"><small><small><small>XXXPOSITIONXXX<br>
      </small></small></small></td>
      <td style="text-align: center; vertical-align: middle; width: 5%;"><small><small><small>XXXASSAYIDXXX<br>
      </small></small></small></td>
      <td
 style="text-align: center; vertical-align: middle; width: 10px;"><small><small><small>XXXASSAYTITLEXXX<br>
      </small></small></small></td>
      <td
 style="text-align: center; vertical-align: middle; width: 10%;"><small><small><small>XXXCHIPIDXXX</small></small></small></td>
      <td style="text-align: center; vertical-align: middle; width: 5%;"><small><small><small>XXXWELLIDXXX<br>
      </small></small></small></td>
      <td
 style="text-align: center; vertical-align: middle; width: 25%;"><small><small><small>XXXTESTREFXXX<br>
      </small></small></small></td>
      <td
 style="width: 40%; vertical-align: middle; text-align: center;"><small><small><small>XXXRESULTXXX<br>
      </small></small></small></td>
    </tr>'''

    foot = "</tbody></table>"

    pagebreak = "<center style='page-break-after: always'></center>"

    texto = ''
    texto += header

    
    i=0

    for sample in samples:

        logging.info("Añadiendo " + " ".join(sample))
        
        if len(sample)==4:  wellid = sample[-4]
        else:               wellid = "-"

        texto += "<tr>"
        texto += "<td style='text-align: center; vertical-align: middle; width: 5%;'><small><small><small>\n"
        texto += sample[-3]
        texto += "\n</small></small></small></td>"
        texto += "<td style='text-align: center; vertical-align: middle; width: 5%;'><small><small><small>\n"
        texto += str(sample[-2][-5:])
        texto += "\n</small></small></small></td>"
        texto += "<td style='text-align: center; vertical-align: middle; width: 10px;'><small><small><small>\n"
        texto += "XXX" + str(sample[-2][-5:]) + "TITLEXXX"
        texto += "\n</small></small></small></td>"
        texto += "<td style='text-align: center; vertical-align: middle; width: 10%;'><small><small><small>\n"
        texto += str(sample[-2])
        texto += "\n</small></small></small></td>"
        texto += "<td style='text-align: center; vertical-align: middle; width: 5%;'><small><small><small>\n"
        texto += wellid
        texto += "\n</small></small></small></td>"
        texto += "<td style='vertical-align: middle; width: 25%;'><small><small><small>\n"
        texto += sample[-1]
        texto += "\n</small></small></small></td>"
        texto += "<td style='width: 40%; vertical-align: middle;'><small><small><small>\n"
        texto += "<!-- " + str(sample[-2][-5:]) + sample[-3] + " -->&nbsp;"
        texto += "\n</small></small></small></td>"
        texto += "</tr>"

        i += 1
        if i%32==0 and i< len(samples):

                texto += foot + pagebreak + header

    texto += foot

    return texto


def CreaWorkingList(template,datastring,runid,runname):

    data = datastring.split(",")

    plateid = data[0]

    number  = data[1]

    j = 2
    samples = []

    for i in xrange(int(number)):

	if int(plateid):

	   samples.append([data[j],data[j+1],data[j+2],data[j+3]])
	   j += 4

	else:

	   samples.append([data[j],data[j+1],data[j+2]])
	   j += 3
    


    dict = {    'XXXRUNIDXXX':    runname + " (" + runid + ")",

                'XXXDATETIMEXXX': time.ctime(),

                '<!-- DATA -->':     CreaFilasMuestras(samples)
            }


        
    return Remplaza(template,dict)




def CheckAssays(datastring):

    ok = 1

    data = datastring.split(",")

    logging.info("Comprobando que están todos los softwares...")

    plateid = data[0]

    number  = data[1]

    j = 2
    samples = []

    for i in xrange(int(number)):

	if int(plateid):

	   samples.append([data[j],data[j+1],data[j+2],data[j+3]])
	   j += 4

	else:

	   samples.append([data[j],data[j+1],data[j+2]])
	   j += 3

    dictAssays = {}
    
    for s in samples:

            dictAssays[s[-2][-5:]] = 1

    for a in dictAssays.keys():

            if not os.path.isdir(
                    os.path.join(
                            os.path.splitdrive(LeeRegistro("InstallDir"))[0] + os.sep,
                            "Genomica","Data","assays",a)):     ok = 0

            #if not os.path.isdir("C:\\genomica\\data\\assays\\" + a): ok = 0

    logging.info({0: "Faltan!!", 1: "Están todos"}[ok])

    return ok

            

            



    
###################################################################################
#
#   Sensovation's main
#
###################################################################################




def main(OutputPath, lang, RunID, Data):

    # Esta función produce el report del workinglist a partir de los datos de las muestras que recibe como
    # argumento
    #
    # La función 1 si ha ido bien o 0 si ha habido problemas
    #
 

    try:

        logging.basicConfig(	level=logging.DEBUG,
	      	            	format='%(asctime)s %(levelname)s %(message)s',
				filename=OutputPath + "\\workinglist.log.txt",
				filemode='a')

        logging.info("Empezamos main...")



        # Arreglo todos los argumentos...
      
        global WorkingListAssayPath
        global idioma

        WorkingListAssayPath = os.path.join(
                os.path.splitdrive(LeeRegistro("InstallDir"))[0] + os.sep,
                "Genomica","Data","scripts") + os.sep
        #WorkingListAssayPath = "C:\\Genomica\\Data\\scripts\\"

       
        idioma               = str(lang)
        RunID                = str(RunID)
        OutputPath           = OutputPath.replace(".\\","\\")
        OutputPath           = OutputPath.replace("\\\\","\\")
        OutputPathNames      = OutputPath.split("\\")

        Runname              = OutputPathNames[-1]
        if not Runname:      Runname = OutputPathNames[-2]


        logging.info("OutputPath " + str(OutputPath))
        logging.info("lang " + idioma)
        logging.info("RunID " + RunID)
        logging.info("Data " + str(Data))



        # Defino el resto de las variables que me hacen falta...        

        wlerror     = WorkingListAssayPath + "workinglist.error.html"
        wltemplate  = WorkingListAssayPath + "workinglist.tmp.html"
        wlreport    = OutputPath + "workinglist.html"
        runreport   = OutputPath + "runreport.html"
        vocfile     = WorkingListAssayPath + "workinglist.voc." + idioma
        logfile     = OutputPath + "log.txt"
        logo_file   = "logo_cabecera1.gif"
        logo_source = WorkingListAssayPath + logo_file
        logo_target = OutputPath + logo_file


        # Y empiezo...

        LeeVoc(vocfile)
        dic = {}

        try:
                logging.info(Data.split(",")[1])
                logging.info(float(len(Data.split(","))-2)/(3+int(Data.split(",")[0])))

                if re.compile("[\'\"]").search(Data) or \
                   (Data.find("\\") + 1) or \
                   float(Data.split(",")[1]) - float(len(Data.split(","))-2)/(3+int(Data.split(",")[0])) <> 0:

                        dic['WRKINTERNALERROR'] = 'TYPEERROR'
                        raise Exception

                if not CheckAssays(Data):

                        dic['WRKINTERNALERROR'] = 'ASSAYNOTFOUND'
                        raise Exception
               

                workinglist = CreaWorkingList(Lee(wltemplate),Data,RunID,Runname)

            	wlready     = Traduce(workinglist)

           	Guarda(wlreport, wlready)
           	Guarda(runreport, wlready)
           	
           	shutil.copy(logo_source,logo_target)

        except Exception,e1:

                msg = "\n" + "\n".join(traceback.format_exception(sys.exc_type,sys.exc_value,sys.exc_traceback))

                logging.error("Sucedio un error en la preparación del informe: " + msg)

                Guarda(wlreport,Traduce(Remplaza(Lee(wlerror),dic,'')))
        
        return 1


    except Exception, e2:

        msg = "\n" + "\n".join(traceback.format_exception(sys.exc_type,sys.exc_value,sys.exc_traceback))

        msg = "Fin del workinglist main. Exit = 0. " + msg
        logging.info(msg)

        return 0
	


if __name__ == "__main__":

        logging.basicConfig(	level=logging.DEBUG,
	       	            	format='%(asctime)s %(levelname)s %(message)s',
				filename="log.txt",
				filemode='a')

        logging.info("Empezamos ejecución...")

        sampleATs    = "0,4,A1,12345001040204,Referencia 1,A2,12345002040204,Referencia 2,A3,12345003040204,Referencia 3,A4,12345004040204,Referencia 4"
                      
        samplestrips = "1,8,1,A1,12345001040204,Referencia 1,2,B1,12345001040204,Referencia 5,3,C1,12345001040204,Referencia 6,4,D1,12345001040204,Referencia 7,5,E1,12345001040204,Referencia 8,6,F1,12345001040204,Referencia 9,7,G1,12345001040204,Referencia 10,8,H1,12345001040204,Referencia 11"

    
        OutputPath   = "C:\\Genomica\\Data\\Results\\2008_3_4_17_58_48_54\\"
        RunID        = "Mi Run ID"
        lang         = "1033"
        samples      = samplestrips


        print main(OutputPath,lang,RunID,samples)



