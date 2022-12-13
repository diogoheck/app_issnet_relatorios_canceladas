import time
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import ElementNotInteractableException


def lancamento_excecao(funcao, driver, *args, **kwargs):
    for i in range(kwargs.get('segundos', 60)):
        try:
            funcao(driver, *args)
            break

        except ElementNotInteractableException as e:
            print('Retry in 1 second', e, funcao.__name__)
            time.sleep(1)

        except NoSuchElementException as e:
            print('Retry in 1 second', e, funcao.__name__)
            time.sleep(1)


def lancamento_excecao_telas(funcao):
    i = 1

    while(not funcao()):
        print(f'Retry in {i} second -> exception window {funcao.__name__}')
        if i == 60:
            break
        time.sleep(1)
        i = i + 1
