from zipfile import ZipFile
import glob
import os
import time

# PASTA_XML = r'U:\ISS\xml_contratados'
PASTA_XML = r'C:\xml'


def extrair_arquivos():

    py_files = glob.glob(PASTA_XML + '/*.zip')
    for arquivo in py_files:
        z = ZipFile(arquivo, 'r')
        z.extractall(PASTA_XML)
        # z.extract(caminho_arquivo)
        z.close()


def remover_arquivos_zip():

    py_files = glob.glob(PASTA_XML + '/*.zip')

    for py_file in py_files:
        try:
            os.remove(py_file)
        except OSError as e:
            print(f"Error:{ e.strerror}")


if __name__ == '__main__':
    extrair_arquivos()
    remover_arquivos_zip()
    time.sleep(10)
