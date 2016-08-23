# -*- coding: cp1252 -*-


from iBuscador import *
from TestSet import *
import logging




class GenAppBuscador(iBuscador):


    def __init__(self):

        self.testSet = None 


    def start(self,inputData):

        self.location = inputData

        
    def dameTests(self):

        ts = TestSet(self.location)
        ts.buildSetLastOnly()

        return ts.getTests()



            
