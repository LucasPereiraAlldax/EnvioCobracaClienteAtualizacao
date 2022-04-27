import shutil
import datetime
import PyPDF2
from bases.consulta_banco import CONSULTA_BANCO
import getpass
import os
import os.path
from tqdm import tqdm


class RECIBOS_PDF:
    user = getpass.getuser()
    # Aqui pegamos a pasta "dir_recibos" e em "dir_empresas" direcionamos ela para .NFe e Boletos Faturamento.
    dir_recibos = fr'C:\Users\{user}\Downloads\recibo_servico'
    dir_empresas = fr'T:\DEPARTAMENTOS\FINANCEIRO\.NFe e Boletos Faturamento - Robo AIA'
    # Aqui pegamos o mês anterior.
    data_hj = datetime.date.today()
    p_mes = data_hj.replace(day=1)
    mes_anterior = p_mes - datetime.timedelta(days=1)
    ano = mes_anterior.year
    mes_anterior = datetime.date.strftime(mes_anterior, '%m%Y')

    # LISTA DE CPF QUE É PARA DESCONSIDERAR
    desconsiderar_cpf = ['63713551472']

    dic_cpf = {'72144831149': 713, '03962016104': 714}
    # '72144831149', '03962016104'

    def acessando_pdf(self):
        # A pasta "recibo_servico" será listada.
        recibos = os.listdir(self.dir_recibos)
        #
        for recibo in recibos:
            # Abrir em pdf
            file = open(fr"{self.dir_recibos}\{recibo}", 'rb')
            read_pdf = PyPDF2.PdfFileReader(file)
            page = read_pdf.getPage(0)
            # Extrair o texto da pagina.
            page_content = page.extractText()
            # Encontra o CNPJ do cliente.
            parte_cliente = page_content.find('Cliente:')
            # Aqui trocamos pontos por vazio.
            cliente_dados = page_content[parte_cliente:]

            if cliente_dados.find('CNPJ:') >= 0:
                part_cnpj = cliente_dados.find('CNPJ:')
                cnpj_dados = cliente_dados[part_cnpj:part_cnpj+24]
                cnpj_dados = cnpj_dados.replace('.', '')
                cnpj_dados = cnpj_dados.replace('-', '')
                cnpj_dados = cnpj_dados.replace('/', '')
                cnpj_dados = cnpj_dados.replace('CNPJ: ', '')

            else:
                part_cnpj = cliente_dados.find('CPF:')
                cnpj_dados = cliente_dados[part_cnpj:part_cnpj + 19]
                cnpj_dados = cnpj_dados.replace('.', '')
                cnpj_dados = cnpj_dados.replace('-', '')
                cnpj_dados = cnpj_dados.replace('/', '')
                cnpj_dados = cnpj_dados.replace('CPF: ', '')

            file.close()
            # Se não tiver o CNPJ passa.
            if cnpj_dados in self.desconsiderar_cpf:
                pass

            elif cnpj_dados == '72144831149' or cnpj_dados == '03962016104':
                codigo = self.dic_cpf[cnpj_dados]
                self.movendo_recibo(str(codigo), recibo)

            else:
                db = CONSULTA_BANCO()
                codigo = db.consultar_cnpj(cnpj_dados)

                self.movendo_recibo(str(codigo), recibo)

    def movendo_recibo(self, codigo, recibo):
        print(f"MOVENDO ARQUIVO - {recibo}")
        pastas_empresas = os.listdir(self.dir_empresas)

        for empresa in pastas_empresas:
            sep_emp = empresa.split('-')

            if sep_emp[0] == codigo:
                if os.path.exists(fr"{self.dir_empresas}\{empresa}\{self.ano}\{self.mes_anterior}\{recibo}"):
                    os.remove(fr"{self.dir_empresas}\{empresa}\{self.ano}\{self.mes_anterior}\{recibo}")

                    shutil.move(fr"{self.dir_recibos}\{recibo}",
                                fr"{self.dir_empresas}\{empresa}\{self.ano}\{self.mes_anterior}")

                else:
                    shutil.move(fr"{self.dir_recibos}\{recibo}",
                                fr"{self.dir_empresas}\{empresa}\{self.ano}\{self.mes_anterior}")

                break


class NF_PDF:
    user = getpass.getuser()
    # Aqui pegamos a pasta "dir_recibos" e em "dir_empresas" direcionamos ela para .NFe e Boletos Faturamento.
    dir_nf = fr'C:\Users\{user}\Downloads\DANFE'
    dir_empresas = fr'T:\DEPARTAMENTOS\FINANCEIRO\.NFe e Boletos Faturamento - Robo AIA'
    # Aqui pegamos o mês anterior.
    data_hj = datetime.date.today()
    p_mes = data_hj.replace(day=1)
    mes_anterior = p_mes - datetime.timedelta(days=1)
    ano = mes_anterior.year
    mes_anterior = datetime.date.strftime(mes_anterior, '%m%Y')

    def acessando_pdf(self):
        lista_nf = tqdm(os.listdir(self.dir_nf))

        for nota_fiscal in lista_nf:
            # Abrir em pdf
            file = open(fr"{self.dir_nf}\{nota_fiscal}", 'rb')
            read_pdf = PyPDF2.PdfFileReader(file)
            page = read_pdf.getPage(0)
            # Extrair o texto da pagina.
            page_content = page.extractText()

            n_nf = page_content.find('SAÍDA1Nº')

            v_nf = page_content[n_nf:n_nf + 14]
            v_nf = v_nf.replace('.', '')
            v_nf = v_nf.replace('SAÍDA1Nº ', '')

            cnpj_cpf = page_content.find('CNPJ / CPF')

            valor_cnpj_cpf = page_content[cnpj_cpf:cnpj_cpf + 28]

            valor_cnpj_cpf = valor_cnpj_cpf.replace('.', '')
            valor_cnpj_cpf = valor_cnpj_cpf.replace('-', '')
            valor_cnpj_cpf = valor_cnpj_cpf.replace('/', '')
            valor_cnpj_cpf = valor_cnpj_cpf.replace('CNPJ  CPF', '')

            file.close()

            db = CONSULTA_BANCO()
            codigo = db.consultar_cnpj(valor_cnpj_cpf)

            if codigo == 727:
                codigo = 573

            self.movendo_nf(str(codigo), nota_fiscal, v_nf)

    def movendo_nf(self, codigo, nota_fiscal, n_nf):
        pastas_empresas = tqdm(os.listdir(self.dir_empresas))

        for empresa in pastas_empresas:
            sep_emp = empresa.split('-')

            if sep_emp[0] == codigo:
                if os.path.exists(fr"{self.dir_empresas}\{empresa}\{self.ano}\{self.mes_anterior}\nota_fiscal_"
                                  fr"{n_nf}.pdf"):
                    os.remove(fr"{self.dir_empresas}\{empresa}\{self.ano}\{self.mes_anterior}\nota_fiscal_{n_nf}.pdf")

                    shutil.move(fr"{self.dir_nf}\{nota_fiscal}",
                                fr"{self.dir_empresas}\{empresa}\{self.ano}\{self.mes_anterior}")

                    os.rename(fr"{self.dir_empresas}\{empresa}\{self.ano}\{self.mes_anterior}\{nota_fiscal}",
                              fr"{self.dir_empresas}\{empresa}\{self.ano}\{self.mes_anterior}\nota_fiscal_"
                              fr"{n_nf}.pdf")

                else:
                    shutil.move(fr"{self.dir_nf}\{nota_fiscal}",
                                fr"{self.dir_empresas}\{empresa}\{self.ano}\{self.mes_anterior}")

                    os.rename(fr"{self.dir_empresas}\{empresa}\{self.ano}\{self.mes_anterior}\{nota_fiscal}",
                              fr"{self.dir_empresas}\{empresa}\{self.ano}\{self.mes_anterior}\nota_fiscal_"
                              fr"{n_nf}.pdf")
                break


class BOLETOS:
    user = getpass.getuser()
    # Aqui pegamos a pasta "dir_recibos" e em "dir_empresas" direcionamos ela para .NFe e Boletos Faturamento.
    dir_boleto = fr'C:\Users\{user}\Downloads'
    dir_empresas = fr'T:\DEPARTAMENTOS\FINANCEIRO\.NFe e Boletos Faturamento - Robo AIA'
    # Aqui pegamos o mês anterior.
    data_hj = datetime.date.today()
    p_mes = data_hj.replace(day=1)
    mes_anterior = p_mes - datetime.timedelta(days=1)
    ano = mes_anterior.year
    mes_anterior = datetime.date.strftime(mes_anterior, '%m%Y')

    # LISTA DE CPF QUE É PARA DESCONSIDERAR
    desconsiderar_cpf = ['63713551472']  # Colocar cnpj das empresas em exceção.

    dic_cpf = {'72144831149': 713, '03962016104': 714}

    def visualizando_arquivo_pdf(self):
        arquivos = os.listdir(self.dir_boleto)
        for arquivo in arquivos:

            if arquivo.find('ALLDAX_') >= 0:
                pos_cli = arquivo.find('cli_')
                pos_doc = arquivo.find('doc_')

                cnpj_cpf = arquivo[pos_cli:]
                cnpj_cpf = cnpj_cpf.replace('cli_', '')
                cnpj_cpf = cnpj_cpf.replace('.pdf', '')
                print(cnpj_cpf)

                n_doc = arquivo[pos_doc:pos_cli]
                n_doc = n_doc.replace('doc_', '')
                n_doc = n_doc.replace('_', '')

                # Se não tiver o CNPJ passa.
                if cnpj_cpf in self.desconsiderar_cpf: # aqui
                    os.remove(fr"{self.dir_boleto}\{arquivo}")
                    pass

                elif cnpj_cpf == '72144831149' or cnpj_cpf == '03962016104':
                    codigo = self.dic_cpf[cnpj_cpf]
                    self.movendo_boleto(str(codigo), arquivo, str(n_doc))

                else:
                    db = CONSULTA_BANCO()
                    codigo = db.consultar_cnpj(cnpj_cpf)

                    if codigo == 727:
                        codigo = 573

                    self.movendo_boleto(str(codigo), arquivo, str(n_doc))

    def movendo_boleto(self, codigo, boleto, n_doc):
        pastas_empresas = tqdm(os.listdir(self.dir_empresas))

        for empresa in pastas_empresas:
            sep_emp = empresa.split('-')

            if sep_emp[0] == codigo:

                shutil.move(fr"{self.dir_boleto}\{boleto}",
                            fr"{self.dir_empresas}\{empresa}\{self.ano}\{self.mes_anterior}")

                os.rename(fr"{self.dir_empresas}\{empresa}\{self.ano}\{self.mes_anterior}\{boleto}",
                          fr"{self.dir_empresas}\{empresa}\{self.ano}\{self.mes_anterior}\boleto_doc_"
                          fr"{n_doc}.pdf")

                break


if __name__ == '__main__':
    rr = BOLETOS()
    rr.visualizando_arquivo_pdf()
