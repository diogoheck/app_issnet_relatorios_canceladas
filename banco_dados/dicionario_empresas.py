from openpyxl import load_workbook


def planilha():

    wb = load_workbook(
        'R:\Compartilhado\Fiscal\lista_clientes_iss\empresas_nova.xlsx')

    ws = wb.active

    return ws


def criar_dicionario_empresas(plan):
    dic = {}
    lista = []
    cabecalho = True
    for linha in plan:
        if not cabecalho:
            if linha[7].value != 'X':
                nova_lista = []
                nova_lista.append(linha[0].value)
                nova_lista.append(linha[1].value)
                nova_lista.append(linha[2].value)
                nova_lista.append(linha[3].value)
                nova_lista.append(str(linha[4].value))
                nova_lista.append(linha[5].value)
                nova_lista.append(linha[6].value)
                dic[str(linha[4].value)] = nova_lista
        cabecalho = False
    return dic


if __name__ == '__main__':

    planilha = planilha()
    dic = criar_dicionario_empresas(planilha)
    print(dic)
