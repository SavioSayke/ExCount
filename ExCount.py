#!/usr/bin/env python
# coding: utf-8

import pandas as pd
import os
import time

directory_of_program = os.getcwd()

document = pd.read_excel(r'{}\table.xlsx'.format(directory_of_program))
all_columns = document.head()

def menu_principal():

    with open('Resultado.txt', 'w') as arquivo:

        while True:
            print()
            print(" ---------------------------------Bem-vindo(a)!---------------------------------")
            print(" | Este programa foi criado com a finalidade de contabilizar quantos elementos |")
            print(" |             possuí em uma determinada coluna de um arquivo excel            |")
            print(" -------------------------------------------------------------------------------")
            print()
            print(" Iniciar - 1")
            print(" Tutorial - 2")
            print(" Créditos - 3")
            print(" Sair - 4")

            menu = input(" Selecione a opção desejada e aperte Enter: ")

            if menu == '1':
                print()
                print(" ------------------------------INICIANDO PROGRAMA------------------------------")
                print()
                quant_columns = input(" Digite a quantidade de colunas que deseja verificar: ")
                total_columns = (document.shape[1] - 1)
                cont = 0

                if quant_columns.isdigit() == False or (int(quant_columns)) < 0 or total_columns < (int(quant_columns)):
                    print(" Quantidade de coluna(s) inexistente(s)!")
                else:
                    for i in range((int(quant_columns))):
                        name_columns = input(" Digite o EXATO nome da coluna que você deseja verificar: ")
                        if name_columns not in all_columns:
                            print()
                            print(" ERROR. Essa coluna não existe, por favor, verifique seu documento e insina o nome válido de uma coluna!")
                            name_columns = input(" Digite o EXATO nome da coluna que você deseja verificar: ")
                        name_element = input(" Digite o EXATO nome do elemento que você desejar pesquisar: ")
                        lista = document[name_columns].tolist()
                        for j in lista:
                            if name_element.isdigit():
                                if j == (int(name_element)):
                                    cont = (cont + 1)
                            else:
                                if j == name_element:
                                    cont = (cont + 1)
                        print("O elemento {} apareceu {} vezes na coluna {}".format(name_element, cont, name_columns), file=arquivo)
                        cont = 0
                        print()
                    load()
                    load_2()
                    load_3()
                    print()
                    print()
                    print(" SUCESSO, o(s) resultado(s) foram salvos nesta pasta como: Resultado.txt")
                obgd()

            elif menu == '2':
                print()
                print(" -----------------------------------TUTORIAL-----------------------------------")
                print(" Passo 1 - Copie seu arquivo Excel para dentro da mesma pasta deste programa")
                print(" Passo 2 - Renomeie seu arquivo Excel EXATAMENTE para: table.xlsx")
                print()
                voltar = input(" Retornar ao menu principal - Digite R e pressione Enter...")
                if voltar == 'R'.upper():
                    menu_principal()

            elif menu == '3':
                crdt()

            elif menu == '4':
                obgd()

            else:
                print()
                print(" ERROR. Selecione uma opção existente do menu!")

def crdt():
    print()
    print(" ------------------------------PROGRAMA CRIADO POR------------------------------")
    print()
    print(" *******************************************************************************")
    print(" *                                  SavioSayke                                 *")
    print(" *******************************************************************************")
    print()
    print(" -------------------------------------------------------------------------------")
    print()
    retornar = input(" Pressione R e Enter para voltar ao menu principal...")
    while True:
        if retornar.upper() == 'R':
            menu_principal()
        else:
            retornar = input(" Pressione R e Enter para voltar ao menu principal...")

def obgd():
    print()
    print(" -----------------------------------------------------------------------------")
    print(" Obrigado por usar este prgrama, espero ter ajudado! :)")
    print(" -----------------------------------------------------------------------------")
    time.sleep(10)
    exit()

def carregando(done):
    print("\r Carregando: {0:50s} {1:.1f}%".format('▊' * int(done * 50), done * 100),end='')

def buscando(done):
    print("\r Buscando elementos: {0:50s} {1:.1f}%".format('▊' * int(done * 50), done * 100),end='')

def terminando(done):
    print("\r Obtendo resultados: {0:50s} {1:.1f}%".format('▊' * int(done * 50), done * 100),end='')

def load():
    for n in range(101):
        carregando(n/100)
        time.sleep(0.05)

def load_2():
    for n in range(101):
        buscando(n/100)
        time.sleep(0.03)

def load_3():
    for n in range(101):
        terminando(n/100)
        time.sleep(0.01)

menu_principal()