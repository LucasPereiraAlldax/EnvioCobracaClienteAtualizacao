import pyodbc


class CONSULTA_BANCO:
    # Conectando o banco da dominio.
    banco = pyodbc.connect('DSN=Contabil')
    cursor = banco.cursor()

    def consultar_cnpj(self, cnpj):
        # Estamos acessando o banco e pegando o CNPJ das empresas.
        self.cursor.execute(
            f"SELECT codi_emp FROM externo.bethadba.geempre WHERE cgce_emp={cnpj}"
        )
        # O codigo = self.cursor.fetchone() nos trará apenas a linha do CNPJ.
        codigo = self.cursor.fetchone()

        return codigo[0]


