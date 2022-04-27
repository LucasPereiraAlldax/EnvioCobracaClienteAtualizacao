import os
import pyodbc
import datetime


class GERAR_PASTAS:
    banco = pyodbc.connect('DSN=Contabil')
    cursor = banco.cursor()

    caminho_pasta = 'T:\DEPARTAMENTOS\FINANCEIRO\.NFe e Boletos Faturamento - Robo AIA'

    def acessando_empresas(self):
        self.cursor.execute(
            f"SELECT codi_emp, nome_emp FROM externo.bethadba.geempre WHERE stat_emp = 'A'"
        )
        info_empresas = self.cursor.fetchall()

        return info_empresas

    def criar_pastas(self):

        lista_empresas = self.acessando_empresas()

        for empresa in lista_empresas:
            codigo_empresa = empresa[0]
            nome_empresa = empresa[1]

            os.mkdir(fr"{self.caminho_pasta}\{codigo_empresa}-{nome_empresa.replace('/', '')}")

    def criar_pastas_anos(self):
        lista_pastas = os.listdir(self.caminho_pasta)

        for pasta in lista_pastas:
            print(pasta)
            for num in range(1, 13):
                data = datetime.date(2021, num, 1)
                data_mod = f"{data:%m%Y}"

                os.mkdir(fr"{self.caminho_pasta}\{pasta}\2021\{data_mod}")

            for num in range(1, 13):
                data = datetime.date(2022, num, 1)
                data_mod = f"{data:%m%Y}"

                os.mkdir(fr"{self.caminho_pasta}\{pasta}\2022\{data_mod}")



gp = GERAR_PASTAS()
gp.criar_pastas_anos()
