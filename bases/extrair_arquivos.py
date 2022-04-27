from zipfile import ZipFile
import os
import getpass
import shutil


class Extraindo:
    # getpass.getuser() serve para referenciar o usuario da maquina.
    user = getpass.getuser()
    # Listagem das "Pastas" que serão excluidas.
    nome_arq_excluir = ['cancelar_nfse_resposta', 'cancelar_nfse_solicitacao', 'consultar_nfse_resposta',
                        'consultar_situacao_resposta', 'demonstrativo', 'distribuicao', 'enviar_rps_resposta',
                        'enviar_rps_solicitacao']
    nome_pasta_excluir = ['XML']

    def extrair_recibos(self, nome_arquivo):
        # Os 3 'z', Descompactam a pasta e encaminha todas elas para o download.
        z = ZipFile(fr'C:\Users\{self.user}\Downloads\{nome_arquivo}', 'r')
        z.extractall(fr'C:\Users\{self.user}\Downloads')
        z.close()

        # excluir pasta .zip
        os.remove(fr'C:\Users\{self.user}\Downloads\{nome_arquivo}')

        # lista_arquivos = os.listdir -> Vai lista o caminho.
        lista_arquivos = os.listdir(rf'C:\Users\{self.user}\Downloads')

        for arquivo in lista_arquivos:
            # Se o lista_arquivos que contem as pastas, estiver em nome_arq_excluir, será apagada com "shutil.rmtree".
            if arquivo in self.nome_arq_excluir:
                shutil.rmtree(rf'C:\Users\{self.user}\Downloads\{arquivo}')

    def extrair_nfe(self, nome_arquivo):
        z = ZipFile(fr'C:\Users\{self.user}\Downloads\{nome_arquivo}', 'r')
        z.extractall(fr'C:\Users\{self.user}\Downloads')
        z.close()

        # excluir pasta .zip
        os.remove(fr'C:\Users\{self.user}\Downloads\{nome_arquivo}')

        lista_arquivos = os.listdir(rf'C:\Users\{self.user}\Downloads')

        for arquivo in lista_arquivos:

            if arquivo in self.nome_pasta_excluir:
                shutil.rmtree(rf'C:\Users\{self.user}\Downloads\{arquivo}')


# pp = Extraindo()
# pp.extrair_pastas('XML')
