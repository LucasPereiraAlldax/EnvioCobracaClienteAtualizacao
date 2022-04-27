from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import os
import datetime
import pyautogui


class onvio:
    def __init__(self):
        self.driver = webdriver.Firefox()
        self.url = "https://onvio.com.br/login/#/"

        self.empresas = os.listdir(fr'T:\DEPARTAMENTOS\FINANCEIRO\.NFe e Boletos Faturamento - Robo AIA')
        # Lista temporaria para fazer o resto das empresas e não precisar apagar.
        self.lista_temp = []
        self.lista_empresas = fr'T:\DEPARTAMENTOS\FINANCEIRO\.NFe e Boletos Faturamento - Robo AIA'

        data_hj = datetime.date.today()
        p_mes = data_hj.replace(day=1)
        mes_anterior = p_mes - datetime.timedelta(days=1)
        self.ano = mes_anterior.year
        self.mes_anterior = datetime.date.strftime(mes_anterior, '%m%Y')

        # Data vencimento completa.
        self.data_de_hoje_vencimento = datetime.date.today()
        self.mes_vencimento = datetime.date.strftime(self.data_de_hoje_vencimento, '%m/%Y')
        print(self.mes_vencimento)
        # Dias respectivos de cada empresa.
        self.emp_diferente_dia_15 = ["692", "693", "695", "574", "698", "700", "706", "707", "709", "737", "901", "713"]
        self.emp_diferente_dia_23 = ["744"]
        self.emp_diferente_dia_25 = ["742"]
        self.emp_diferente_dia_28 = ["500"]  # Lembrar que fevereiro tem 28 dias aí o dia 30 vai dar problema.
        self.emp_diferente_dia_30 = ["712", "943", "888", "900", "769", "828", "964", "965", "963", "962"]

        self.nome_mes = {'01': "Janeiro", "02": "Fevereiro", "03": "Março", "04": "Abril", "05": "Maio", "06": "Junho",
                         "07": "Julho", "08": "Agosto", "09": "Setembro", "10": "Outubro", "11": "Novembro",
                         "12": "Dezembro"}

        # Empresas que no campo de procura é preciso colocar o nome da empresa ao inves do ID.
        self.emp_que_nao_aceitam_enter = ["187", "43", "461", "54", "658", "673", "7", "700", "717", "777",
                                          "850", "93", "965"]

        # Todas essas empresas foram revisadas como cortesia.
        self.emp_cortesia = ["211", "229", "236", "248", "249", "252", "260", "301", "355", "363", "376", "455", "556",
                             "574", "612", "666", "7", "702", "717", "721", "758", "777", "778", "781", "798", "813",
                             "818", "830", "837", "859", "873", "883", "887", "897", "904", "905", "906", "938",
                             "939", "943", "946", "172", "956", "957", "958", "959", "960", "961", "574", "967", "936"]

    def acessando_site(self):
        self.driver.get(self.url)
        time.sleep(5)

    def fazendo_loguin(self):
        self.driver.find_element(By.XPATH, '//*[@id="trta1-auth"]/label[1]/span[2]/input').send_keys("aia@alldax.com")
        time.sleep(2)
        self.driver.find_element(By.XPATH, '//*[@id="trta1-auth"]/label[2]/span[2]/input').send_keys("All@12345")
        time.sleep(2)
        self.driver.find_element(By.XPATH, '//*[@id="trta1-auth"]/div/button').click()
        time.sleep(2)
        self.driver.find_element(By.XPATH, '//*[@id="trta1-mfs-later"]').click()
        time.sleep(20)
        self.driver.find_element(By.XPATH, '/html/body/bm-optional-header/bm-staff-nav/bm-nav/nav/ul/li[2]/a').click()
        time.sleep(5)

    def check_vazio_ou_cheio(self, empresa):
        pasta_empresa = fr"{self.lista_empresas}\{empresa}\{self.ano}\{self.mes_anterior}"
        verificar_arquivo = os.listdir(pasta_empresa)
        # verificar_arquivo = os.listdir(self.lista_empresas)

        if len(verificar_arquivo) == 0:
            existe = 0
        else:
            existe = 1

        return existe

    def acessando_pasta_financeiro(self):
        global existe_pasta_financeiro
        time.sleep(5)
        items_pastas = self.driver.find_elements(By.TAG_NAME, "a")
        for itens in items_pastas:
            item = itens.text
            if item == 'Financeiro':
                existe_pasta_financeiro = 1
                itens.click()
                break
            else:
                existe_pasta_financeiro = 0

        if existe_pasta_financeiro == 0:
            print("CRIAR PASTA FINANCEIRO")
            time.sleep(5)
            self.driver.find_element(By.XPATH, '//*[@id="dms-fe-legacy-components-'
                                               'client-documents-new-menu-button"]').click()
            time.sleep(2)
            self.driver.find_element(By.XPATH, '//*[@id="dms-fe-legacy-components-client-documents-'
                                               'new-folder-button"]').click()
            time.sleep(2)
            self.driver.find_element(By.ID, 'containerName').send_keys('Financeiro')
            time.sleep(2)
            self.driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div/div[3]/button[1]').click()

            time.sleep(5)
            items_pastas = self.driver.find_elements(By.TAG_NAME, "a")
            for itens in items_pastas:
                item = itens.text
                if item == 'Financeiro':
                    itens.click()
                    break

    def acessando_pasta_honorarios(self):
        global existe_pasta_honorarios
        time.sleep(5)
        items_financeiros = self.driver.find_elements(By.TAG_NAME, 'a')
        for pasta in items_financeiros:
            honorarios = pasta.text
            if honorarios == 'Honorários' or honorarios == 'Honorarios':
                existe_pasta_honorarios = 1
                pasta.click()
                break
            else:
                existe_pasta_honorarios = 0

        if existe_pasta_honorarios == 0:
            print("CRIAR PASTA HONORARIOS")

            time.sleep(5)
            self.driver.find_element(By.XPATH, '//*[@id="dms-fe-legacy-components-'
                                               'client-documents-new-menu-button"]').click()
            time.sleep(2)
            self.driver.find_element(By.XPATH, '//*[@id="dms-fe-legacy-components-client-documents-'
                                               'new-folder-button"]').click()
            time.sleep(2)

            self.driver.find_element(By.ID, 'containerName').send_keys('Honorarios')

            self.driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div/div[3]/button[1]').click()

            time.sleep(5)
            items_financeiros = self.driver.find_elements(By.TAG_NAME, 'a')
            for pasta in items_financeiros:
                honorarios = pasta.text
                if honorarios == 'Honorarios' or honorarios == 'Honorários':
                    pasta.click()
                    break

    def acessando_pasta_ano(self):
        global existe_pasta_ano
        time.sleep(5)
        items_anos = self.driver.find_elements(By.TAG_NAME, 'a')
        for ano in items_anos:
            ano_comp = ano.text
            if ano_comp == str(self.ano):
                existe_pasta_ano = 1
                ano.click()
                break
            else:
                existe_pasta_ano = 0

        if existe_pasta_ano == 0:
            print("CRIAR PASTA ANO")
            time.sleep(5)
            self.driver.find_element(By.XPATH, '//*[@id="dms-fe-legacy-components-'
                                               'client-documents-new-menu-button"]').click()
            time.sleep(2)
            self.driver.find_element(By.XPATH, '//*[@id="dms-fe-legacy-components-client-documents-'
                                               'new-folder-button"]').click()
            time.sleep(2)

            self.driver.find_element(By.ID, 'containerName').send_keys(self.ano)

            self.driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div/div[3]/button[1]').click()
            time.sleep(5)
            items_anos = self.driver.find_elements(By.TAG_NAME, 'a')
            for ano in items_anos:
                ano_comp = ano.text
                if ano_comp == str(self.ano):
                    ano.click()
                    break

    def listando_empresas(self):
        global existe_pasta_financeiro, existe_pasta_honorarios, existe_pasta_ano
        empresas = os.listdir(self.lista_empresas)

        for empresa in empresas:
            print(empresa)
            id_empresas = empresa.split('-')
            cod_emp = id_empresas[0]
            nome_empresa = id_empresas[1]

            existe = self.check_vazio_ou_cheio(empresa)
            if existe == 0:
                print('NÃO TEM ARQUIVO')
                pass

            else:
                print("POSSUI ARQUIVO PARA O ONVIO")
                if cod_emp in self.emp_cortesia:
                    print('Empresa cortesia')
                    pass
                else:
                    time.sleep(2)
                    self.consulta_emp(cod_emp, nome_empresa)
                    time.sleep(2)
                    pyautogui.press('enter')
                    time.sleep(5)

                    try:
                        self.acessando_pasta_financeiro()

                    except:
                        self.driver.refresh()
                        time.sleep(30)
                        print('#Pesou o site, Mais 30 seg financeiro')
                        self.acessando_pasta_financeiro()

                    try:
                        self.acessando_pasta_honorarios()

                    except:
                        self.driver.refresh()
                        time.sleep(30)
                        print('#Pesou o site, Mais 30 seg honorarios')
                        self.acessando_pasta_honorarios()

                    try:
                        self.acessando_pasta_ano()

                    except:
                        self.driver.refresh()
                        time.sleep(30)
                        print('#Pesou o site, Mais 30 seg ano')
                        self.acessando_pasta_ano()

                    try:
                        self.acessando_pasta_mes()

                    except:
                        self.driver.refresh()
                        time.sleep(30)
                        print('#Pesou o site, Mais 30 seg mês')
                        self.acessando_pasta_mes()

                    try:
                        self.processo_upl(empresa, cod_emp)

                    except:
                        self.driver.refresh()
                        time.sleep(30)
                        print('#Pesou o site, Mais 30 seg uploud')
                        self.processo_upl(empresa, cod_emp)

                    # try:
                    #     self.vencimento(cod_emp)
                    #
                    # except:
                    #     self.driver.refresh()
                    #     time.sleep(30)
                    #     print('#Pesou o site, Mais 30 seg vencimento')
                    #     self.vencimento(cod_emp)

                    # Clica no campo de busaca, e limpa.
                    time.sleep(5)
                    self.driver.find_element(By.XPATH, '/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div['
                                                       '1]/div/div[1]/navbar-left/dms-clients-combobox/div/div['
                                                       '2]/div[1]/input').click()
                    time.sleep(5)
                    print('limpando')
                    self.driver.find_element(By.XPATH,
                                             '/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[1]/div/div[1]/'
                                             'navbar-left/dms-clients-combobox/div/div[2]/div[1]/input').click()
                    pyautogui.hotkey('backspace')

    def processo_upl(self, empresa, cod_emp):
        time.sleep(10)
        pasta_empresa = fr"{self.lista_empresas}\{empresa}\{self.ano}\{self.mes_anterior}"
        verificar_arquivo = os.listdir(pasta_empresa)
        for arquivo in verificar_arquivo:
            time.sleep(6)
            self.driver.find_element(By.XPATH, '//*[@id="dms-fe-legacy-components-client-'
                                               'documents-upload-button"]').click()
            time.sleep(3)
            pyautogui.typewrite(fr'{pasta_empresa}\{arquivo}')
            time.sleep(5)
            salvar = pyautogui.locateOnScreen('img/abrir.PNG', confidence=0.9)
            pyautogui.click(salvar)
            time.sleep(2)
            try:
                self.vencimento(cod_emp)

            except:
                self.driver.refresh()
                time.sleep(30)
                print('#Pesou o site, Mais 30 seg vencimento')
                self.vencimento(cod_emp)

    def consulta_emp(self, codigo_emp, nome_emp):
        if codigo_emp in self.emp_que_nao_aceitam_enter:
            print('acessando pelo nome')
            self.driver.find_element(By.XPATH, '/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[1]/'
                                               'div/div[1]/navbar-left/dms-clients-combobox/div/div[2]/'
                                               'div[1]/input').send_keys(nome_emp)
        else:
            print('acessando pelo codigo')
            self.driver.find_element(By.XPATH, '/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div[1]/'
                                               'div/div[1]/navbar-left/dms-clients-combobox/div/div[2]/'
                                               'div[1]/input').send_keys(codigo_emp)

    def acessando_pasta_mes(self):
        time.sleep(8)
        self.driver.find_element(By.XPATH, '//*[@id="dms-fe-legacy-components-'
                                           'client-documents-new-menu-button"]').click()
        time.sleep(1)
        self.driver.find_element(By.XPATH, '//*[@id="dms-fe-legacy-components-client-documents-'
                                           'new-folder-button"]').click()
        time.sleep(1)
        escrever_nome_pasta = f"{self.nome_mes[self.mes_anterior[:2]]}-" \
                              f"{self.ano}"

        self.driver.find_element(By.ID, 'containerName').send_keys(escrever_nome_pasta)

        self.driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div/div[3]/button[1]').click()

        time.sleep(7)
        pasta_comp = self.driver.find_elements(By.TAG_NAME, 'a')
        for comp in pasta_comp:
            comp_text = comp.text
            if comp_text == escrever_nome_pasta:
                time.sleep(3)
                comp.click()
                break

    def vencimento(self, cod_emp):
        time.sleep(5)
        self.driver.find_element(By.XPATH, '/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div['
                                           '2]/div/section/div/documents-detail-pane/div/div/dms'
                                           '-document-grid/div/div/div[14]/div[3]/div/div/div/div/i').click()
        time.sleep(4)
        self.driver.find_element(By.XPATH,
                                 '//*[@id="dms-fe-legacy-components-client-documents-manage-docs'
                                 '-menu-button"]').click()
        time.sleep(4)
        self.driver.find_element(By.XPATH, '/html/body/bm-main/main/div[1]/ui-view/div[2]/div/div['
                                           '2]/div/section/div/documents-detail-pane/div/dms-document'
                                           '-grid-toolbar/dms-toolbar/div/ul/li[4]/ul/li['
                                           '11]/a/i').click()
        time.sleep(4)
        self.driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/form/div[1]/input').click()
        time.sleep(3)
        pyautogui.press('tab')
        time.sleep(2)
        print(cod_emp)
        if cod_emp in self.emp_diferente_dia_15:
            pyautogui.write(f'15/{self.mes_vencimento}')
            time.sleep(2)
            self.driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[3]/button[1]').click()

        elif cod_emp in self.emp_diferente_dia_23:
            pyautogui.write(f'23/{self.mes_vencimento}')
            time.sleep(2)
            self.driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[3]/button[1]').click()

        elif cod_emp in self.emp_diferente_dia_25:
            pyautogui.write(f'25/{self.mes_vencimento}')
            time.sleep(2)
            self.driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[3]/button[1]').click()

        elif cod_emp in self.emp_diferente_dia_28:
            pyautogui.write(f'28/{self.mes_vencimento}')
            time.sleep(2)
            self.driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[3]/button[1]').click()

        # elif cod_emp in self.emp_diferente_dia_30:
        #     pyautogui.write(f'30/{self.mes_vencimento}')
        #     time.sleep(2)
        #     self.driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[3]/button[1]').click()

        else:
            pyautogui.write(f'20/{self.mes_vencimento}')
            time.sleep(2)
            self.driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[3]/button[1]').click()


