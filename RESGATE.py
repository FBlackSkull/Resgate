        ###############################################################
        #	Target:		Resgate de Arquivos                           #
        #	Company:	Prodata Mobility Brasil                       #
        #	Author:		Fernando Soares & Caio Teixeira               #
        ###############################################################


###########################################################################################################
################################# IMPORTANDO DEMAIS PARAMETROS NECESSÁRIOS ################################
###########################################################################################################
import os, time, sys
import json, threading, re, shutil, tarfile, subprocess, platform, paramiko, socket
from paramiko import SSHClient, AutoAddPolicy, transport, SecurityOptions, ssh_exception
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from termcolor import colored
from getpass import getpass

###########################################################################################################
############################################ DATAS PARA SCRIPT ############################################
###########################################################################################################
DATA = datetime.today().strftime('-%d-%m-%y')
DATAXLSX = datetime.today().strftime('%d/%m/%Y')
ANO = datetime.today().strftime('%Y')
RTC = datetime.today().strftime('%m%d%H%M%y')

###########################################################################################################
############################################### CABEÇALHO #################################################
###########################################################################################################
print(colored('######################### RESGATE DE ARQUIVOS - V3 BYONIC #########################\n', 'blue', 'on_white', attrs=['bold']))

print('                                _---------_ ')
print('                               |  _______  |')
print('                               | |       | |')
print('                               | |P      | |')
print('                               | |   M   | |')
print('                               | |      B| |')
print('                               | |_______| |')
print('                               \  --   --  /')
print('                                \  (( ))  / ')
print('                                 \_______/  ')
print()
print(colored(' ________  ________  ________  ________  ________  _________  ________ ', 'blue', attrs=['bold']))
print(colored('|\   __  \|\   __  \|\   __  \|\   ___ \|\   __  \|\___   ___\\\   __  \ ', 'blue', attrs=['bold']))
print(colored('\ \  \|\  \ \  \|\  \ \  \|\  \ \  \_|\ \ \  \|\  \|___ \  \_\ \  \|\  \ ', 'blue', attrs=['bold']))
print(colored(' \ \   ____\ \   _  _\ \  \\ \  \ \  \ \\\ \ \   __  \   \ \  \ \ \   __  \ ', 'blue', attrs=['bold']))
print(colored('  \ \  \___|\ \  \\\  \\\ \  \\ \  \ \  \_\\\ \ \  \ \  \   \ \  \ \\ \  \ \  \ ', 'blue', attrs=['bold']))
print(colored('   \ \__\    \ \__\\\ _\\\ \_______\ \_______\ \__\ \__\   \ \__\ \ \__\ \__\ ', 'blue', attrs=['bold']))
print(colored('    \|__|     \|__|\|__|\|_______|\|_______|\|__|\|__|    \|__|  \|__|\|__| ', 'blue', attrs=['bold']))


###########################################################################################################
####################################### INICIANDO RESGATE #################################################
###########################################################################################################
print('\n                         ', end='')
print(colored(' --> INCIANDO RESGATE <-- \n', attrs=['underline', 'bold']))

# EMPRESA
g_EMPRESA = input(colored("NOME DA EMPRESA? ", 'cyan', attrs=['bold'])).upper().strip()
# OS
g_OS = input(colored("NUMERO DA OS? ", 'cyan', attrs=['bold'])).strip()
# NUMERO DE SERIE
g_NS = input(colored("NUMERO DE SERIE? ", 'cyan', attrs=['bold'])).strip()
# SELEÇÃO SPTRANS OU PROJETO (TOP)
g_SP_CMT = input(colored("\nTIPO DE EQUIPAMENTO:\nDIGITE 1 PARA SPTRANS:\nDIGITE 2 PARA CMT-PROJETOS:\nDIGITE 3 PARA TOP-BOM:\nOPÇÃO: ", 'cyan', attrs=['bold']))
print()

### VERIFICAR INFORMAÇÕES ###
print(colored('VERIFIQUE AS INFORMAÇÕES!', 'yellow'))
INFORMACOES = input("Digite R para CORRIGIR ou ENTER para Continuar...\n",).upper().strip()
if INFORMACOES == 'R':
    print('REINICIANDO...')
    os.system('python RESGATE.py')
elif INFORMACOES == 'RR':
    print('REINICIANDO...')
    os.system('python RESGATE.py')
else:
    pass

###########################################################################################################
########################################### SELECIONANDO PROJETO ##########################################
###########################################################################################################
if g_SP_CMT == "1":
        g_SP_CMT2 = "SPTRANS"
elif g_SP_CMT == "2":
        g_SP_CMT2 = "CMT-PROJETOS"
elif g_SP_CMT == "3":
        g_SP_CMT2 = "TOP-BOM"
else:
    print(colored('     ---> NENHUMA DAS OPÇÕES FOI SELECIONADA <---', 'red', attrs=['bold']))
    ########## VERIFICAR SE REINICIA OU NAO ##########
    print(colored("\n\nREINICIAR?", 'yellow'))
    REINICAR_PROJETO = input('Digite S para SIM ou ENTER para Fechar...\n').upper().strip()
    if REINICAR_PROJETO == 'S':
        print('REINICIANDO...')
        os.system('python RESGATE.py')
    elif REINICAR_PROJETO == 'SIM':
        print('REINICIANDO...')
        os.system('python RESGATE.py')
    elif REINICAR_PROJETO == 'SS':
        print('REINICIANDO...')
        os.system('python RESGATE.py')    
    else:
        print('FECHANDO...\n')
        time.sleep(2)
        os.system('TASKKILL /IM python.exe /T /F')

print('Equipamento do tipo: '+g_SP_CMT2+'\n')
