import pandas as pd
import os
from scripts.Formatacao import Formatacao
formatacao = Formatacao()

class VerificarNotaBaixada:
    def verificadorNotaBaixada(lerPlanilha):
        lerPlanilha = pd.read_excel(formatacao.get_caminhoNotas(), formatacao.get_sheetnameNotas(), dtype={'Documento': str})
        lerPlanilha['Nome'] = lerPlanilha['Nome'].astype(str)

        cont = 0
        cont_baixada = 0

        for x in range(int(formatacao.quantidadeNotas())):
            
            cont += 1
            nome_cliente = lerPlanilha['Nome'][x]

            caminho_pdf =f'C:\\Users\\Suporte\\OneDrive\\Área de Trabalho\\Notas fiscais\\ANO 2024\\08 - AGOSTO 2024\\{nome_cliente}.pdf'

            if os.path.exists(caminho_pdf):
                cont_baixada += 1
                print(f"       {cont} - ✅ JÁ EXISTE DOWNLOAD PARA ESSE CLIENTE {nome_cliente}")
                lerPlanilha.drop(index=[x], inplace=True)
            else:
                print(f"       {cont} - ❌ NÃO EXISTE DOWNLOAD PARA ESSE CLIENTE {nome_cliente}")
                
        if cont_baixada > 0:
            print('\n')
            input(f"Alerta! Você possui {cont_baixada} notas já baixadas!\nAPERTE [Enter] PARA RETIRAR DA PLANILHA ")
        
        return lerPlanilha


