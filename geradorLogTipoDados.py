from itertools import combinations_with_replacement
from numpy import NaN
import pandas as pd
import os
from pathlib import Path
# import funcAuxTemplate
import win32com.client as win32

# Função para Checar se o arquivo existe no diretório


def checarArquivo(valor):
    arquivo = Path(valor)
    if arquivo.is_file():
        return valor
    else:
        os.system("cls")
        novoValor = str(
            input("O arquivo não existe, digite um novo arquivo: \n")).strip()
        novoValor = novoValor + ".xlsx"
        return checarArquivo(novoValor)

# Função para checar o tamanho da String.


def checarTamanhoString(valStr):
    valStr = str(valStr)
    tamanhoString = len(valStr)
    if tamanhoString > 1500:
        return True
    return False

def checarCaracterEsp(val):
    val = str(val)
    if "*" in val:
        return True
    return False

# Checa se os headers existem.


def checkHeaders(digitadoC, HeadOri):
    nw = []
    for dados in digitadoC:
        cont = 0
        if dados not in HeadOri:
            cont += 1
            valorNovo = str(
                input("Valores de Headers não existem , digite novamente por favor!! \n"))
            valorNovo = valorNovo.split(";")
            return checkHeaders(valorNovo, HeadOri)
    if cont > 0:
        return valorNovo
    else:
        return digitadoC


def checkDelimitador(value):
    cont = 0
    if ";" not in value:
        cont += 1
        valorNovo = str(
            input("String não contém o delimitador ';', favor digite novamente!! \n"))
        return checkDelimitador(valorNovo)
    if cont > 0:
        return valorNovo
    else:
        return value


def checkDelimitadorT(value):
    cont = 0
    if len(value) > 1:
        cont += 1
        valorNovo = str(
            input("String não contém o delimitador ';', favor digite novamente!! \n"))
        return checkDelimitador(valorNovo)
    if cont > 0:
        return valorNovo
    else:
        return value


def checarNumerico(v):
    v = str(v)
    if not v.isdigit():
        return True
    return False

# =========================================== Welcome to the Program! ==========================

infEnt = str(input("Informe o arquivo de entrada:\n"))
infEnt = infEnt.strip()
infEnt = infEnt + ".xlsx"
infEnt = checarArquivo(infEnt)
x = pd.read_excel(infEnt)

# Guarda os nomes dos Headers.
nomeHeader = []

for nomesHeader in x:
    nomeHeader.append(nomesHeader)


# Tipo de dado Texto

# Veriica quais são os headers
checarTiposTexto = str(input("Informe as colunas para o tipo texto:\n"))
checarTiposTexto = checarTiposTexto.strip()
novo = checarTiposTexto.split(";")
novo = checkHeaders(novo, nomeHeader)

for i in novo:
    novosVal = []
    for lendoArq in novo:
        linhaTx = 1
        for c in x[lendoArq]:
            if c is not NaN:
                c = str(c).strip()
                linhaTx += 1
                checkLenght = checarTamanhoString(c)
                if checkLenght:
                    novosVal.append(
                        f"A coluna {lendoArq} contém mais de 1500 caracteres na linha {linhaTx} , verifique!!")
                checl = checarCaracterEsp(c)
                if checl:
                    novosVal.append(
                        f"A coluna {b} contém caracter especial '*' na linha {linhaTx} , verifique!!")


# Tipo de dado Numérico

# Verifica quais são os headers
checarTiposNumerico = str(input("Informe as colunas para o tipo Numérico:\n"))
checarTiposNumerico = checarTiposNumerico.strip()
novoNum = checarTiposNumerico.split(";")
novos = checkHeaders(novoNum, nomeHeader)

for i in novos:
    novosVals = []
    for b in novos:
        linhaTN = 1
        for c in x[b]:
            if c is not NaN:
                c = str(c).strip()
                linhaTN += 1
                checkNumber = checarNumerico(c)
                if checkNumber:
                    novosVals.append(
                        f"A coluna {b} na linha {linhaTN} , não contém somente números , por favor verifique!!")

# Tipo de dado Lista fixo

# Verifica quais são os headers
checarTiposLista = str(input("Informe as colunas para o tipo lista Fixo:\n"))
checarTiposLista = checarTiposLista.strip()
novoLista = checarTiposLista.split(";")
novoLista = checkHeaders(novoLista, nomeHeader)


# Tipo de dado Lista Múltiplo

# Verifica quais são os headers
checarTiposListaM = str(
    input("Informe as colunas para o tipo lista Seleção Multipla:\n"))
checarTiposListaM = checarTiposListaM.strip()
novoListaM = checarTiposLista.split(";")
novoListaM = checkHeaders(novoListaM, nomeHeader)

# Informa o arquivo no qual estão os dados oriundos da Ferramenta.
infPlan = str(
    input("Informe o arquivo , para checagem de valores de Lookup no CRM:\n"))
infPlan = infPlan.strip()
infPlan = infPlan + ".xlsx"
infPlanCheck = checarArquivo(infPlan)
y = pd.read_excel(infPlanCheck)


newVl = []
for i in novoLista:
    for dd in y[i]:
        dd = str(dd)
        newVl.append(dd.strip())
        novosValLista = []
    for b in novoLista:
        linhaL = 1
        for c in x[b]:
            linhaL += 1
            if c is not NaN:
                c = str(c)
                c = c.strip()
                if c not in newVl:
                    novosValLista.append(
                        f"A coluna {b} não contém {c} na linha {linhaL}")


newVlM = []
for i in novoListaM:
    for dd in y[i]:
        dd = str(dd)
        newVlM.append(dd.strip())
    novosValListaM = []
    vm = []
    for b in novoListaM:
        linhaLM = 1
        for c in x[b]:
            linhaLM += 1
            if c is not NaN and c not in newVlM:
                c = str(c)
                c = c.strip()
                if ',' not in c or c is None:

                    dps = c.split(',')
                    for ab in dps:
                        if ab in dps:
                            vm.append(
                                f"A coluna {b} não contém {ab} na linha {linhaLM}")

# Faz o merge com todas as Listas.
tot = vm + novosVal + novosVals

tamanhoArr = len(tot)

if tamanhoArr > 0:
    print("Há alterações a serem feita neste arquivo.")
    valAlt = str(input("Digite o nome do arquivo a ser criado: \n"))
    valAlt = valAlt.strip()
    valAlt = valAlt + ".txt"
    f = open(valAlt, 'w')
    for jj in tot:
        f.write(jj+"\n")
    dads = pd.DataFrame(data=tot)
    removeX = valAlt.replace(".txt", ".xlsx")
    dads.to_excel(removeX, index=False)
    print("Arquivo Gerado com sucesso!!")
    f.close()
    os.system("pause")
else:
    print("Nao há Alterações para serem feitas no arquivo!")
    os.system("pause")

os.system("pause")
