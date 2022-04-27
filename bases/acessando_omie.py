from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import datetime
import pyautogui
from bases.extrair_arquivos import Extraindo
import os
import getpass
from bases.acessando_pdf import RECIBOS_PDF, NF_PDF
from bases.omie_api import Query


class omie:
    def __init__(self):
        self.driver = webdriver.Firefox()
        self.url = "https://app.omie.com.br/my-apps/"
        self.data = '03/03/2022'  # Colocar da data que o Eduardo Subir o faturamento.
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

        self.aplicativos = ['ALLDAX ASSESSORIA E CONTABILIDADE LTDA', 'ALLDAX CONSULTORIA E TECNOLOGIA']

        data_hoje = datetime.date.today()
        self.data_inicio = datetime.date(data_hoje.year, data_hoje.month, 4) # lembrando que deveria ser do 1 ao 4
        self.data_fim = datetime.date(data_hoje.year, data_hoje.month, 5)

        self.user = getpass.getuser()

        self.empresas_excecoes = ["551", "179", "488", "499", "512", "540", "712"]

    def acessando_site(self):
        self.driver.get(self.url)
        time.sleep(5)

    def fazendo_loguin(self):
        # Acessando OME.
        self.driver.find_element(By.NAME, "email").send_keys("aia@alldax.com")
        self.driver.find_element(By.ID, "btn-continue").click()
        time.sleep(5)
        self.driver.find_element(By.NAME, "current-password").send_keys("alldax1234")
        self.driver.find_element(By.ID, "btn-login").click()
        time.sleep(6)

    def acessando_empresa(self):
        # Na Alldax Assessoria teremos apenas recibo/boleto

        acessos_aplicativos = ['ALLDAX ASSESSORIA E CONTABILIDADE LTDA', 'ALLDAX CONSULTORIA E TECNOLOGIA']
        for acessos in acessos_aplicativos:
            aplicativos = self.driver.find_elements(By.TAG_NAME, "a")
            lista_aplicativos = []
            for app in aplicativos:
                texto_aplicativos = app.text
                lista_aplicativos.append(texto_aplicativos)
                if acessos in lista_aplicativos:
                    app.click()
                    # Aqui farei todo processo apos o clique no app. Lembrando que fará alldax 2 depois a 3
                    time.sleep(15)
                    self.acessando_painel_nfse()
                    #  self.acessando_painel_nfse() está pronto agora deve passar para o boleto.
                    # proximo passo boleto nfse
                    print('para aquiiiii')
                    # Quando chegar no boleto ferei uma condição caso uma muda o app key caso outro muda tbm.
                    break

    def acessando_painel_nfse(self):
        # Acessando Serviços e NFS-e
        global nome_zip
        # self.driver.find_element(By.XPATH, "/html/body/div[2]/div[4]/div[2]/ul/li[3]/a/img").click()
        # time.sleep(2)
        self.driver.find_element(By.XPATH, "/html/body/div[2]/header/div[2]/ul/li[2]").click()
        time.sleep(7)

        # 1. Buscando Painel de NFS-e, notas fiscais eletronicas e recibos de serviços emitidos.
        self.driver.find_element(By.XPATH, "/html/body/div[2]/div[1]/div[7]/div/div[1]/div[1]/span/input"). \
            send_keys("Painel de NFS-e")
        time.sleep(2)
        self.driver.find_element(By.XPATH, "/html/body/div[2]/div[1]/div[7]/div/div[1]/div[2]/button").click()
        time.sleep(7)
        self.driver.find_element(By.XPATH, "/html/body/div[2]/div[1]/div[7]/div/div[2]/div[4]"
                                           "/div[2]/div/a[2]/h4/div/span").click()
        # Adicionando data no filtro
        time.sleep(5)
        self.driver.find_element(By.XPATH, '//*[@id="d50625c26"]').send_keys(f"{self.data_inicio:%d/%m/%Y}")
        self.driver.find_element(By.XPATH, '//*[@id="d50625c27"]').send_keys(f"{self.data_fim:%d/%m/%Y}")
        time.sleep(1)
        self.driver.find_element(By.XPATH, '/html/body/div[2]/div[6]/div[2]/div[3]/div[3]/div/div[2]/button[1]'
                                           '/span[2]').click()

        # Exportar XMLs e recibos
        time.sleep(3)
        xmls = self.driver.find_elements(By.TAG_NAME, 'a')
        for xml in xmls:
            exportar = xml.text
            if exportar.find('Exportar XMLs e Recibos') >= 0:
                xml.click()
        time.sleep(5)
        self.driver.find_element(By.XPATH, '/html/body/ul[2]/li/div/div[2]/button[1]').click()
        time.sleep(3)
        self.driver.find_element(By.XPATH, '/html/body/ul[2]/li/div/div[2]/button[1]').click()
        # time.sleep(90)

        while True:
            janela = pyautogui.locateOnScreen('img/janela_down.png', confidence=0.9)

            if janela is not None:
                break

        salvar = pyautogui.locateOnScreen('img/salvar_arquivo.png', confidence=0.9, region=janela)
        pyautogui.click(salvar)
        time.sleep(2)
        pyautogui.hotkey('enter')

        # Extraindo recibo_servico do zip
        time.sleep(15)
        arquivos = os.listdir(rf'C:\Users\{self.user}\Downloads')
        for arquivo in arquivos:
            if arquivo.find('nfse_') >= 0:
                nome_zip = arquivo
                break

        # extrair = Extraindo()
        # extrair.extrair_recibos(nome_zip) #############################################

        # Chamando a classe que move os recibos para respectivas pastas das empresas
        # rp = RECIBOS_PDF()
        # rp.acessando_pdf()
        time.sleep(5)

    # 2. Painel de NF-e, notas fiscais eletronicas emitidas e recebidas.
    def acessando_painel_nfe(self):
        global nome_zip  # estava comentado o self de baixo.
        self.driver.find_element(By.XPATH, '/html/body/div[2]/div[6]/div[1]/button/i').click()
        time.sleep(4)
        self.driver.find_element(By.XPATH, "/html/body/div[2]/header/div[2]/ul/li[2]").click()
        time.sleep(4)
        self.driver.find_element(By.XPATH, "/html/body/div[2]/div[1]/div[7]/div/div[1]/div[1]/span/input").clear()
        time.sleep(4)
        self.driver.find_element(By.XPATH, "/html/body/div[2]/div[1]/div[7]"
                                           "/div/div[1]/div[1]/span/input").send_keys("Painel de NF-e")
        time.sleep(4)
        self.driver.find_element(By.XPATH, "/html/body/div[2]/div[1]/div[7]/div/div[1]/div[2]/button").click()
        time.sleep(5)
        self.driver.find_element(By.XPATH, "/html/body/div[2]/div[1]/div[7]/div/"
                                           "div[2]/div[4]/div[2]/div/a[3]/h4/div/span").click()
        time.sleep(5)
        self.driver.find_element(By.XPATH, '//*[@id="d50629c20"]').send_keys(f"{self.data_inicio:%d/%m/%Y}")
        self.driver.find_element(By.XPATH, '//*[@id="d50629c25"]').send_keys(f"{self.data_fim:%d/%m/%Y}")
        time.sleep(4)
        self.driver.find_element(By.XPATH, "/html/body/div[2]/div[6]/div[2]/div[3]/"
                                           "div[3]/div/div[1]/button/span[2]").click()
        time.sleep(4)
        self.driver.find_element(By.XPATH, "/html/body/div[2]/div[6]/div[2]/div[3]/div[1]/ul/a[1]/div[3]").click()
        time.sleep(7)
        self.driver.find_element(By.XPATH, "/html/body/ul[3]/li/div/div[2]/button[1]").click()

        while True:
            janela = pyautogui.locateOnScreen('img/janela_down.png', confidence=0.9)

            if janela is not None:
                break

        salvar = pyautogui.locateOnScreen('img/salvar_arquivo.png', confidence=0.9, region=janela)
        pyautogui.click(salvar)
        time.sleep(2)
        pyautogui.hotkey('enter')

        time.sleep(15)
        arquivos = os.listdir(rf'C:\Users\{self.user}\Downloads')
        for arquivo in arquivos:
            if arquivo.find('omie_') >= 0:
                nome_zip = arquivo
                break

        # extrair = Extraindo()
        # extrair.extrair_nfe(nome_zip) ################################################

        nf = NF_PDF()
        nf.acessando_pdf()

    def acessando_boletos(self):
        time.sleep(7)
        # self.driver.find_element(By.XPATH, "/html/body/div[2]/div[6]/div[1]/button").click()
        time.sleep(2)
        self.driver.find_element(By.XPATH, "/html/body/div[2]/header/div[2]/ul/li[2]").click()
        time.sleep(2)
        self.driver.find_element(By.XPATH, "/html/body/div[2]/div[1]/div[7]/div/div[1]/div[1]/span/input").clear()
        self.driver.find_element(By.XPATH, "/html/body/div[2]/div[1]/div[7]/" "div/div[1]/div[1]/span/input"). \
            send_keys("exibir todas as contas a receber")
        time.sleep(1)
        self.driver.find_element(By.XPATH, "/html/body/div[2]/div[1]/div[7]/div/div[1]/div[2]/button").click()
        time.sleep(5)
        self.driver.find_element(By.XPATH, "/html/body/div[2]/div[1]/div[7]/div/"
                                           "div[2]/div[4]/div[2]/div/a[1]/h4/div/span").click()
        time.sleep(15)

        # LOOP PAGINAS
        while True:
            time.sleep(10)
            num_pag = self.driver.find_element(By.CLASS_NAME, 'ui-iggrid-pagelinkcurrent').text
            # TRAMITES DA PAGINAS
            elementos_boleto = self.driver.find_elements(By.TAG_NAME, 'td')
            elementos_boleto_tr = self.driver.find_elements(By.TAG_NAME, 'tr')

            for tr in elementos_boleto_tr:
                valor_tr = tr.text
                var1 = valor_tr.split('-')
                var2 = var1[0].split(' ')
                if len(var2) <= 1:
                    pass
                else:
                    result = var2[-2]  # erro
                    for boleto in elementos_boleto:
                        nome = boleto.text
                        if result not in self.empresas_excecoes and nome.find("A vencer") >= 0:
                            boleto.click()
                            time.sleep(1)
                            self.driver.find_element(By.XPATH,
                                                     "/html/body/div[2]/div[6]/div[2]/div[3]/div[1]/ul/a[5]").click()

                            time.sleep(8)
                            self.driver.find_element(By.XPATH,
                                                     "/html/body/div[2]/div[8]/div[2]/div[3]/div[3]/div/div[2]/a[3]") \
                                .click()

                            time.sleep(5)
                            self.driver.switch_to.window(self.driver.window_handles[-1])
                            time.sleep(2)
                            self.driver.find_element(By.ID, 'download').click()

                            time.sleep(4)
                            pyautogui.hotkey('enter')

                            self.driver.close()

                            self.driver.switch_to.window(self.driver.window_handles[-1])

                            # movendo boleto para pasta da empresa
                            # bb = BOLETOS()
                            # bb.visualizando_arquivo_pdf()

                            # botao voltar
                            self.driver.find_element(By.XPATH, '/html/body/div[2]/div[8]/div[1]/button').click()
                            time.sleep(3)

                    time.sleep(10)
                    self.driver.find_element(By.XPATH,
                                             "/html/body/div[2]/div[6]/div[2]/div[3]/div[2]/div[2]/div/div[61]/div"
                                             "/div[3]").click()

                    nova_pag = self.driver.find_element(By.CLASS_NAME, 'ui-iggrid-pagelinkcurrent').text

                    if num_pag == nova_pag:
                        print("PAGINA CONTINUOU NA MESMA")
                        break

            self.driver.close()

# entrar = omie()
# entrar.acessando_site()
# entrar.fazendo_loguin()
# entrar.acessando_empresa()
