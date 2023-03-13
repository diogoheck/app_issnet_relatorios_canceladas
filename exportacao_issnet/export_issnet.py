# -*- coding: utf-8 -*-
# -*- coding: cp1252 -*-
# import webdriver
from __future__ import unicode_literals
from re import X
import time
from tkinter import E
from selenium.webdriver.support.select import Select
import pyautogui
from exception.lancar_excecao import lancamento_excecao, lancamento_excecao_telas
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import ElementNotInteractableException
from bs4 import BeautifulSoup
from selenium.webdriver.common.by import By

PASTA_RELATORIO_CANCELADAS = r'C:\ISS\relatorio_canceladas'
PASTA_LIVRO_PRESTADOS = r'C:\ISS\livro_prestados'
# PASTA_RELATORIO_CANCELADAS = r'C:\ISS\relatorio_contratados'
# PASTA_LIVRO_PRESTADOS = r'C:\ISS\livro_prestados'


def inserir_IE(driver, IE):
    driver.find_element(By.XPATH, '//*[@id="txtCae"]').send_keys(IE)


def clicar_botao_procurar_empresa(driver):
    driver.find_element(By.CLASS_NAME, 'nc-search').click()


def seleciona_empresa(driver, IE):
    lancamento_excecao(inserir_IE, driver, IE)
    lancamento_excecao(clicar_botao_procurar_empresa, driver)


def menu_toggle(driver):
    driver.find_element(By.XPATH, '//*[@id="menu-toggle"]').click()


def mudar_barra_lateral(driver):
    for _ in range(10):
        try:
            elemento = driver.find_element(By.ID, 'wrapper')
            break
        except ElementNotInteractableException as e:
            print('Retry in 1 second', e)
            time.sleep(1)

    if elemento.get_attribute("class") != 'toggled':
        lancamento_excecao(menu_toggle, driver)


def seleciona_menu(driver):
    lancamento_excecao(mudar_barra_lateral, driver)


def seleciona_menu_nfe(driver):
    driver.find_element(By.XPATH,
                        '//*[@id="Menu1_MenuPrincipal"]/ul/li[8]/div/span[3]').click()


def selecionar_menu_nota_eletronica(driver, IE, empresa):
    lancamento_excecao(seleciona_menu_nfe, driver)


def consulta_nfe(driver):
    driver.find_element(By.XPATH,
                        '//*[@id="Menu1_MenuPrincipal"]/ul/li[8]/ul/li[2]/div/a').click()


def consulta_nota_eletronica(driver, simples, empresa):
    lancamento_excecao(consulta_nfe, driver)


def consulta_cancelamento_exc(driver, intervalo_nf):
    driver.find_element(By.XPATH,
                        '//*[@id="Menu1_MenuPrincipal"]/ul/li[8]/ul/li[3]/div/a').click()


def consultar_solicitacao_cancelamento(driver, intervalo_nf):
    lancamento_excecao(consulta_cancelamento_exc, driver, intervalo_nf)


# mudar_frame
def switch_frame(driver):
    frame = driver.find_element(By.XPATH,
                                '//*[@id="iframe"]')
    driver.switch_to.frame(frame)


def mudar_frame(driver):
    lancamento_excecao(switch_frame, driver)


def switch_relatorio(driver):
    frame = driver.find_element(By.XPATH,
                                '//*[@id="viewer"]')
    driver.switch_to.frame(frame)


def mudar_para_relatorio(driver):
    lancamento_excecao(switch_relatorio, driver)


def frame_principal(driver):
    driver.switch_to.default_content()


def mudar_frame_principal(driver):
    lancamento_excecao(frame_principal, driver)


def clicar_na_serie_nota(driver):
    select = Select(driver.find_element(By.ID, 'ddlSerie'))
    select.select_by_value('22')


def clicar_na_status_solicitacao(driver):
    select = Select(driver.find_element(By.ID, 'ddlStatusSolicitacao'))
    select.select_by_value('0')


def selecionar_serie_nota(driver):
    lancamento_excecao(clicar_na_serie_nota, driver)


def clicar_filtros_adicionais(driver):
    driver.find_element(By.ID, "imbArrow").click()


def selecionar_filtros_adicionais(driver):
    lancamento_excecao(clicar_filtros_adicionais, driver)


def inserir_data_inicial(driver, dt_inicial):
    driver.find_element(By.XPATH,
                        '//*[@id="txtDtEmissaoIni"]').send_keys(dt_inicial)


def selecionar_data_inicial(driver, dt_inicial):
    lancamento_excecao(inserir_data_inicial, driver, dt_inicial)


def inserir_data_final(driver, dt_final):
    driver.find_element(By.XPATH,
                        '//*[@id="txtDtEmissaoFim"]').send_keys(dt_final)


def selecionar_data_final(driver, dt_final):
    lancamento_excecao(inserir_data_final, driver, dt_final)


def clicar_buscar_notas(driver):
    driver.find_element(By.XPATH,
                        '//*[@id="btnLocalizar2"]/span').click()


def buscar_notas(driver):
    lancamento_excecao(clicar_buscar_notas, driver)


def clica_entrar_empresa(driver):
    driver.find_element(By.XPATH,
                        '//*[@id="lblNomeEmpresa"]').click()


def trocar_empresa(driver):
    lancamento_excecao(clica_entrar_empresa, driver)


def clicar_botao_imprimir(driver):
    driver.find_element(By.XPATH,
                        '//*[@id="btnImprimir"]/span').click()


def carregar_tela_impressao():
    return pyautogui.locateOnScreen('reportManager.png')


def carregar_tela_salvar():
    return pyautogui.locateOnScreen('botao_salvar_EN.png')


def carregar_tela_print():
    return pyautogui.locateOnScreen('botao_salvar_EN.png')


def carregar_janela_imprimir():
    return pyautogui.locateOnScreen('botao_imprimir.png')


def carregar_imprimir_final():
    return pyautogui.locateOnScreen('botao_salvar_final.png')


def carregar_tela_nenhum_registro(driver):
    try:
        driver.find_element(By.XPATH,
                            '/html/body/div[1]/div/div/div[3]/div/button').click()
        return True
    except:
        return False


def gerar_relatorio_canceladas_sem_movim(driver, IE, empresa, caminho,
                                         dt_inicial):

    mes = dt_inicial[3:5]
    ano = dt_inicial[6:10]
    nome_salvar = f'{PASTA_RELATORIO_CANCELADAS}\\{empresa[4]} {mes}{ano}'
    nome_salvar = nome_salvar + '.txt'
    print('passei aqui xxxxxxxxxxxxxxxxxxx')
    with open(nome_salvar, 'a') as saida:
        print(f'SEM MOV', file=saida)


def gerar_relatorio_canceladas(driver, IE, empresa, caminho, padrao,
                               dt_inicial):

    mes = dt_inicial[3:5]
    ano = dt_inicial[6:10]
    nome_salvar = f'{PASTA_RELATORIO_CANCELADAS}\\{empresa[4]} {mes}{ano}'
    nome_salvar = nome_salvar + '.txt'
    print(nome_salvar)
    intervalo_nfe = []
    table = driver.find_element(By.TAG_NAME, "table")
    rows = table.find_elements(By.TAG_NAME, "tr")

    if len(rows) > 2:
        lista_notas = list(rows)

        # nota_inicial = lista_notas[1].find_elements_by_tag_name("td")[0].text
        with open(nome_salvar, 'a') as saida:
            lista_notas_sem_rodape = lista_notas[1:-1]

            for tr in lista_notas_sem_rodape:
                print(tr.text)
                print(f'{tr.text}', file=saida, sep=';')
                # for td in tr.find_elements_by_tag_name("td"):
                #     print(f'impressaoooo....{td.text}')
    else:
        with open(nome_salvar, 'a') as saida:
            print('sem movimento', file=saida)

        # nota_final = lista_notas[-2].find_elements_by_tag_name("td")[0].text


def intervalo_nfe_exc(driver):
    return lancamento_excecao(pegar_intervalo_notas_mes, driver)


def pegar_rodape(driver, inicial):
    table = driver.find_element(By.TAG_NAME, "table")
    rows = table.find_elements(By.TAG_NAME, "tr")
    lista_notas = list(rows)
    tam_lista = len(lista_notas)
    rodape = lista_notas[tam_lista -
                         1].find_elements(By.TAG_NAME, "td")[0].text
    dic = {}
    print(f'este rodape: {rodape}')
    novo_rodape = rodape.split(' ')
    for rod in novo_rodape:
        dic[rod] = novo_rodape.index(rod)

    return dic


def pega_numero_notas_pagina(driver):
    table = driver.find_element(By.TAG_NAME, "table")
    rows = table.find_elements(By.TAG_NAME, "tr")
    notas = []
    if len(rows) > 2:
        lista_notas = list(rows)
        for nota in lista_notas:
            nota = nota.find_elements(By.TAG_NAME, "td")[0].text
            notas.append(nota)
    return notas


def apenas_1_pagina(driver):
    table = driver.find_element(By.TAG_NAME, "table")
    rows = table.find_elements(By.TAG_NAME, "tr")
    if len(rows) > 2:

        lista_notas = list(rows)
        tam_lista = len(lista_notas)
        rodape = lista_notas[tam_lista -
                             1].find_elements(By.TAG_NAME, "td")[0].text
        if rodape == '1':
            return True
        else:
            return False


def pegar_intervalo_notas_mes(driver):
    intervalo_nfe = []
    table = driver.find_element(By.TAG_NAME, "table")
    rows = table.find_elements(By.TAG_NAME, "tr")

    if len(rows) > 2:

        lista_notas = list(rows)

        nota_inicial = lista_notas[1].find_elements(By.TAG_NAME, "td")[0].text
        tam_lista = len(lista_notas)

        nota_final = lista_notas[tam_lista -
                                 2].find_elements(By.TAG_NAME, "td")[0].text
        rodape = lista_notas[tam_lista -
                             1].find_elements(By.TAG_NAME, "td")[0].text

        print(nota_inicial, nota_final)
        print(rodape.split(' '))

        intervalo_nfe.append(nota_inicial)
        intervalo_nfe.append(nota_final)
    return intervalo_nfe


def inserir_num_inicial_exc(driver, num_inicial):
    driver.find_element(By.XPATH,
                        '//*[@id="txtNumInicial"]').send_keys(num_inicial)


def inserir_num_inicial(driver, num_inicial):
    lancamento_excecao(inserir_num_inicial_exc, driver, num_inicial)


def inserir_num_final_exc(driver, num_inicial):
    driver.find_element(By.XPATH,
                        '//*[@id="txtNumFinal"]').send_keys(num_inicial)


def inserir_num_final(driver, num_inicial):
    lancamento_excecao(inserir_num_final_exc, driver, num_inicial)


def procurar_canceladas_except(driver):
    driver.find_element(By.CLASS_NAME, 'nc-search').click()


def procurar_canceladas(driver):
    lancamento_excecao(procurar_canceladas_except, driver, segundos=1)


def clicar_pagina_except(driver, rodape, segundos=1):
    print(f'cliqueeeeeeeeeee {rodape}')


def clicar_pagina(driver, rodape, segundos=1):
    lancamento_excecao(clicar_pagina_except, driver,
                       rodape, segundos=1)


def pegar_notas_inicio_fim(lista_notas):
    notas_finais = []
    intervalo = []
    for notas in lista_notas:
        notas_finais.append(notas[1])
        notas_finais.append(notas[-2])
    # print(notas_finais)
    notas_finais = sorted(notas_finais)
    intervalo.append(notas_finais[0])
    intervalo.append(notas_finais[-1])
    print(f'intervalo {intervalo}')
    return intervalo


def percorrer_guias_notas_canceladas(driver, IE, empresa, caminho, padrao,
                                     dt_inicial):
    lista_notas = []
    rodape_anterior = ''
    pagina_table = 0
    primeiro_rodape = True
    ultimo_indice = -200
    primeiro_nota = True
    lista_notas = []
    if not carregar_tela_nenhum_registro(driver):
        if not apenas_1_pagina(driver):
            gerar_relatorio_canceladas(
                driver, IE, empresa, caminho, padrao,
                dt_inicial)
            print('herreeee')
            lista_indices_ja_usados = []
            inicio = True
            while (True):
                if inicio:
                    vlr_partida = 1
                    inicio = False
                print('passei aqui')
                rodape_inicial = pegar_rodape(driver, vlr_partida)
                if rodape_inicial == rodape_anterior:
                    break
                else:
                    rodape_anterior = rodape_inicial
                print(rodape_inicial)
                try:
                    for valor, indice in rodape_inicial.items():
                        if not primeiro_rodape:
                            indice += 1
                        if int(indice) != 0 and (valor not in lista_indices_ja_usados):
                            print(int(indice))

                            driver.find_element(By.XPATH,
                                                f'//*[@id="dgDocumentos"]/tbody/tr[12]/td/a[{int(indice)}]').click()

                            gerar_relatorio_canceladas(
                                driver, IE, empresa, caminho, padrao,
                                dt_inicial)
                            lista_indices_ja_usados.append(valor)
                    lista_indices_ja_usados.append(str(indice + 1))
                    # driver.find_element_by_xpath(
                    #     f'//*[@id="dgDocumentos"]/tbody/tr[12]/td/a[{int(indice)}]').click()

                except:
                    break

                primeiro_rodape = False

        else:

            gerar_relatorio_canceladas(
                driver, IE, empresa, caminho, padrao,
                dt_inicial)
    else:
        gerar_relatorio_canceladas(driver, IE, empresa, caminho,
                                   dt_inicial)


def percorrer_menus_servicos_contratados_relatorios(driver, IE, dt_inicial,
                                                    dt_final, simples,
                                                    empresa, caminho):

    intervalo_nfe = []
    seleciona_empresa(driver, IE)  # ok
    time.sleep(1)
    seleciona_menu(driver)  # ok
    time.sleep(1)
    selecionar_menu_nota_eletronica(driver, IE, empresa)  # ok
    time.sleep(1)
    consulta_nota_eletronica(driver, simples, empresa)  # ok
    time.sleep(1)
    mudar_frame(driver)
    time.sleep(0.5)
    # selecionar_serie_nota(driver)
    selecionar_filtros_adicionais(driver)
    time.sleep(0.5)
    selecionar_data_inicial(driver, dt_inicial)
    time.sleep(0.5)
    selecionar_data_final(driver, dt_final)
    time.sleep(0.5)
    buscar_notas(driver)
    pagina_table = 0
    rodape_anterior = ''
    ultimo_indice = -200
    primeiro_nota = True
    lista_notas = []
    primeiro_rodape = True
    if not carregar_tela_nenhum_registro(driver):
        if not apenas_1_pagina(driver):
            lista_notas.append(pegar_intervalo_notas_mes(driver))
            print('passo1')
            lista_indices_ja_usados = []
            inicio = True
            while (True):
                if inicio:
                    vlr_partida = 1
                    inicio = False
                print('passo2')
                rodape_inicial = pegar_rodape(driver, vlr_partida)
                print(
                    f'rodape inicial {rodape_inicial} rodape anterior{rodape_anterior}')
                if rodape_inicial == rodape_anterior:
                    break
                else:
                    rodape_anterior = rodape_inicial
                try:
                    for valor, indice in rodape_inicial.items():
                        if not primeiro_rodape:
                            indice += 1
                        if int(indice) != 0 and (valor not in lista_indices_ja_usados):
                            print(f'indice {int(indice)}')

                            driver.find_element(By.XPATH,
                                                f'//*[@id="dgDocumentos"]/tbody/tr[12]/td/a[{int(indice)}]').click()

                            lista_notas.append(
                                pega_numero_notas_pagina(driver))
                            lista_indices_ja_usados.append(valor)
                    lista_indices_ja_usados.append(str(indice + 1))
                    # driver.find_element_by_xpath(
                    #     f'//*[@id="dgDocumentos"]/tbody/tr[12]/td/a[{int(indice)}]').click()

                except:
                    break
                primeiro_rodape = False
            intervalo_nfe = pegar_notas_inicio_fim(lista_notas)
            print(f'novo intervalo {intervalo_nfe}')

        else:

            intervalo_nfe = pegar_intervalo_notas_mes(driver)

        # print(f'intervalo nfe {intervalo_nfe}')
        mudar_frame_principal(driver)
        time.sleep(1)
        consultar_solicitacao_cancelamento(driver, intervalo_nfe)
        time.sleep(1)
        mudar_frame(driver)
        time.sleep(1)
        clicar_na_status_solicitacao(driver)
        time.sleep(1)
        inserir_num_inicial(driver, intervalo_nfe[0])
        inserir_num_final(driver, intervalo_nfe[1])
        time.sleep(1)
        procurar_canceladas(driver)
        time.sleep(1)
        # mudar_frame_principal(driver)
        # lancamento_excecao(menu_toggle, driver)
        percorrer_guias_notas_canceladas(
            driver, IE, empresa, caminho, 'LIVRO ISSQN CANCELADAS', dt_inicial)


def empresa_do_simples(empresa):
    if empresa[5] == 'N':
        return False
    else:
        return True


def exportar_empresas_contratados(driver, dic_empresas, dt_inicial, dt_final):
    caminho = False
    # print(dic_empresas)
    for empresa in dic_empresas.values():
        Identificador = empresa[4]
        simples = empresa_do_simples(empresa)

        percorrer_menus_servicos_contratados_relatorios(driver, Identificador,
                                                        dt_inicial,
                                                        dt_final, simples,
                                                        empresa, caminho)

        caminho = True
        time.sleep(2)
        # gerar segunda empresa em diante
        mudar_frame_principal(driver)
        trocar_empresa(driver)
        print('*' * 50)
        print('*' * 50)
        print(empresa)
        print('*' * 50)
        print('*' * 50)

    print('Finalizando......', end='')
    for _ in range(10):
        print('..............', end='')
        time.sleep(1)
    driver.close()
    driver.quit()
