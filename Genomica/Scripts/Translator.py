

from xml.dom import minidom
import re


class Translator:

    def __init__(self,vocfile):

        self.voc = {}

        try:

            xmldoc = minidom.parse(vocfile)

            words = xmldoc.getElementsByTagName('word')

            for word in words:

                try:
                    self.voc[word.attributes['key'].value] = word.attributes['value'].value


                except:
                    txt = "Error al leer " + word.toxml()
                    print txt

        except:

            raise Exception("Error al parsear el voc " + vocfile)


    def voz(self,k):

        try:
            return self.voc[k]

        except:
            return k


    def translate(self,source):

        cadenas = []

        if type(source) == type([]):

            cadenas = source

        else:

            cadenas.append(source)

        resultados =[]

        for res in cadenas:

            for k in self.voc.keys():

                tmp = ''
                while tmp <> res:

                        tmp = res
                        res = re.sub ( k, self.voz(k), res, re.MULTILINE | re.DOTALL)
			tmp = res

            resultados.append(res)

        if len(resultados) <> len(cadenas):   raise Exception("Imposible traduccion")

        if len(resultados)==1:     return resultados[0]
        else:                      return resultados



    
