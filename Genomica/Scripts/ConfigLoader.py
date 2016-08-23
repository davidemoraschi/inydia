

from xml.dom import minidom



class ConfigLoader:

    def __init__(self,configFile):

        self.configDoc = minidom.parse(configFile)


    def lee(self,tag):

        rta = {}

        nodes = self.configDoc.getElementsByTagName(tag)

        for att in nodes[0].attributes.keys():

            rta[att] = nodes[0].attributes[att].value

        return rta


