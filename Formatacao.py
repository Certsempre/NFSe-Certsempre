import datetime
import pandas as pd

class Formatacao:
    def __init__(self):
        self._data = datetime.date.today().strftime('%d/%m/%Y')
        self._caminhoNotas = 'C:\\Users\\Suporte\\Downloads\\Cópia de Pix JP Agosto(1).xlsx'
        self.sheetnameNotas = '26 a 28.08'
        
    def get_data(self):
        return self._data

    def quantidadeNotas(self):
        tamanho = lerPlanilha.shape
        num_linhas = tamanho[0]
        return num_linhas
    
    def get_caminhoNotas(self):
        return self._caminhoNotas
    
    def get_sheetnameNotas(self):
        return self.sheetnameNotas

    def get_pegarValor(self):
        self._valor = str(self._valor)
        self._valor += "00"
        self._valor = int(self._valor)
        return self._valor

lerPlanilha = pd.read_excel(Formatacao().get_caminhoNotas(), Formatacao().get_sheetnameNotas(), dtype={'Documento': str})

