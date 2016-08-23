


from iSelector import *



class GenAppSelector(iSelector):

    def __init__(self):

        self.selectedTests = []
    

    def start(self,inputData):

        self.selectedTests = inputData
        

    def dameSeleccionados(self):

        return self.selectedTests
