import requests
import datetime
import json
import webbrowser
import time
import pyautogui
from bases.acessando_pdf import RECIBOS_PDF, NF_PDF, BOLETOS


class Query:
    def __init__(self, data):
        self.data = str(data)
        self.listagem = ['01796430000124', '18374073000109', '18563901000157', '19434611000176', '19434956000120',
                         '19422438000196', '19534902000136', '08880518000179', '12021740000193', '12216789000100',
                         '12482516000107', '22872227000160', '14734352000185', '00057240000122', '24065619000142',
                         '22872227000240', '01363739000120', '30060762000144', '28319160000117', '31118810000170',
                         '33943767000103', '34752596000106', '34773712000165', '34844572000179', '35273891000134',
                         '35561134000166', '35724005000141', '36176687000168', '36546864000150', '27447090000110',
                         '38332832000188', '39495485000177', '39875701000100', '40398101000187', '41085219000118',
                         '41085191000119', '41246946000110', 'w', '89709861115', '42360849000116', '42361015000125',
                         '42724527000109', '42834254000155', '44520121000130', '44504219000101', '44519743000148',
                         '44519768000141', '44504271000150', '44504808000181', '33440014000185']

        # self.corpos_jsons = [
        #     ['''{"call":"PesquisarLancamentos",
        #         "app_key":"1458834874497",
        #         "app_secret":"50d3875c8e0c43b95e0ee34925fa00a4",
        #         "param":[{"nPagina":1,"nRegPorPagina":1,
        #          "cStatus": "AVENCER",
        #          "cNatureza": "R"}]}''',
        #      "1458834874497",
        #      "50d3875c8e0c43b95e0ee34925fa00a4"],
        # ]
        self.corpos_jsons = [
            ['''{"call":"PesquisarLancamentos",
                "app_key":"1458873207792",
                "app_secret":"048dec4d103c7395bf27a6d8a89197f7",
                "param":[{"nPagina":1,"nRegPorPagina":1,
                 "cStatus": "AVENCER",
                 "cNatureza": "R"}]}''',
             "1458834874497",
             "50d3875c8e0c43b95e0ee34925fa00a4"],
        ]

    def requisita(self):
        global link
        url = "https://app.omie.com.br/api/v1/financas/pesquisartitulos/"
        headers = {'Content-Type': 'application/json'}
        # Captura do nCodTitulo para utilização na requisição dos títulos
        for corpos in self.corpos_jsons:
            app_key = corpos[1]
            app_secret = corpos[2]
            response = requests.post(url=url, headers=headers, data=corpos[0])
            contas = response.content

            # Captura das response's de todas as empresas
            response_completa = json.loads(contas)
            # Captura da quantidade de páginas
            paginas = response_completa['nTotPaginas']
            print("A quantidade de páginas é de:", paginas)

            for i in range(1, paginas):
                corpo_paginado = '''{"call":"PesquisarLancamentos","app_key":"''' + app_key + '''","app_secret":"''' + app_secret + '''","param":[{"nPagina":''' + str(
                    i) + ''',"nRegPorPagina":1, "cStatus": "AVENCER", "cNatureza": "R"}]}'''
                response2 = requests.post(url=url, headers=headers, data=corpo_paginado)
                conteudo = response2.content
                titulos = json.loads(conteudo)
                print('oiiii', titulos)
                titulo = titulos['titulosEncontrados']
                nValorTitulo = titulo[0]['cabecTitulo']['nValorTitulo']
                nCodTitulo = titulo[0]['cabecTitulo']['nCodTitulo']
                cStatus = titulo[0]['cabecTitulo']['cStatus']
                cCPFCNPJCliente = titulo[0]['cabecTitulo']['cCPFCNPJCliente']
                dDtVenc = titulo[0]['cabecTitulo']['dDtVenc']
                dDtEmissao = titulo[0]['cabecTitulo']['dDtEmissao']

                if cCPFCNPJCliente in self.listagem:
                    ...
                    print('CNPJ exceção')

                # elif dDtEmissao > self.data:
                #     print('Data maior do que a fornecida, fim dos boletos.')
                #     break

                else:
                    # Verifica o status do boleto
                    url_bol = "https://app.omie.com.br/api/v1/financas/contareceberboleto/"
                    headers_bol = {'Content-Type': 'application/json'}
                    corpo_bol = '''{"call":"ObterBoleto","app_key":"''' + app_key + '''","app_secret":"''' + app_secret + '''","param":[{"nCodTitulo":''' + str(
                        nCodTitulo) + ''',"cCodIntTitulo":""}]}'''

                    response4 = requests.post(url=url_bol, headers=headers_bol, data=corpo_bol)
                    conteudo_bol = response4.content
                    situacao = json.loads(conteudo_bol)

                    confere = str(situacao)

                    if confere.__contains__("Nenhum boleto foi gerado para essa conta a receber"):
                        print("Nenhum boleto foi gerado para essa conta a receber")
                        continue

                    status = situacao['cCodStatus']

                    if status != "0":
                        print("Boleto não gerado para este título")

                    elif status == "0":

                        if dDtEmissao == self.data:
                            # Captura do boleto
                            url_bol = "https://app.omie.com.br/api/v1/financas/pesquisartitulos/"
                            headers_bol = {'Content-Type': 'application/json'}
                            corpo_bol = '''{"call":"ObterURLBoleto","app_key":"''' + app_key + '''","app_secret":"''' + app_secret + '''","param":[{"nCodTitulo":''' + str(
                                nCodTitulo) + ''',"cCodIntTitulo":""}]}'''

                            response3 = requests.post(url=url_bol, headers=headers_bol, data=corpo_bol)
                            conteudo_bol = response3.content
                            boletos = json.loads(conteudo_bol)
                            link = boletos['cLinkBoleto']
                            print(f'Data de Emissão: {dDtEmissao}\nCnpj da empresa: {cCPFCNPJCliente}\nStatus: {cStatus}\nValor do '
                                  f'Documento: {nValorTitulo}\nData de vencimento: {dDtVenc}\nLink do boleto: {link}')

                            webbrowser.open(link)
                            time.sleep(3)
                            # Baixando Boleto
                            baixar = pyautogui.locateOnScreen('img/Dowloader.PNG', confidence=0.9)
                            pyautogui.click(baixar)
                            time.sleep(3)
                            # Salvando Boleto no Downloads
                            pyautogui.hotkey('alt', 'l')
                            time.sleep(2)
                            # Fechando a página do Boleto.
                            pyautogui.hotkey('ctrl', 'w')
                            time.sleep(2)
                            # movendo boleto para pasta da empresa
                            # mover_boleto = BOLETOS()
                            # mover_boleto.visualizando_arquivo_pdf() ##########################


# if __name__ == '__main__':
#     run = Query('15/10/2021')
#     run.requisita()
