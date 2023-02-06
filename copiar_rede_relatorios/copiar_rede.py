# -*- coding: cp1252 -*-

import os
import shutil
PASTA_RELATORIO_CONTRATADOS = r'U:\ISS\relatorio_contratados'
PASTA_LIVRO_PRESTADOS = r'U:\ISS\livro_prestados'
PASTA_CLIENTES = r'O:\Clientes'
PASTA_RELATORIO_CANCELADAS = r'U:\ISS\relatorio_canceladas'


def criar_nova_pasta(caminho):
    os.makedirs(caminho)


def existe_essa_pasta(caminho):
    return os.path.exists(caminho)


def criar_dicionario_empresas():
    with open('empresas_teste.csv', encoding='cp1252') as arquivo:
        dic_empresas = {}
        cabecalho = True
        for registro in arquivo:
            emp = registro.strip().split(';')
            if not cabecalho:
                dic_empresas[emp[4]] = emp
            cabecalho = False
    return dic_empresas


def copiar_arquivos(origem, destino):
    shutil.copyfile(origem, destino)


def mover_arquivo(origem, destino):
    shutil.move(origem, destino)


def contratados(empresas, competencia):

    for diretorio, subpastas, arquivos in os.walk(PASTA_RELATORIO_CANCELADAS):

        for arquivo in arquivos:

            nome_empresa = empresas.get(arquivo.split(' ')[-2])[1]
            # data = arquivo.split(' ')[-1]
            origem = os.path.join(diretorio, arquivo)
            nome_arquivo = origem.split('\\')[-1]
            nova_pasta = f'{PASTA_CLIENTES}\\{nome_empresa}\\Dpto Tributário\\2023\\Declarações\\ISSQN\\'
            nova_pasta = nova_pasta + competencia
            if not existe_essa_pasta(nova_pasta):
                criar_nova_pasta(nova_pasta)
            destino = f'{nova_pasta}\\{nome_arquivo}'
            print('*' * 50)
            print(destino)
            print('*' * 50)
            mover_arquivo(origem, destino)


if __name__ == '__main__':

    dic_empresas = criar_dicionario_empresas()
    contratados(dic_empresas)

    print('SUCESSOOOO...VAI DESCANSAR')
