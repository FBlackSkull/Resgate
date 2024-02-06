        ###############################################################
        #	Target:		Resgate de Arquivos                           #
        #	Company:	Prodata Mobility Brasil                       #
        #	Author:		Fernando Soares & Caio Teixeira               #
        ###############################################################

###########################################################################################################
##################################### INSTALANDO PARAMETROS PRIMARIOS #####################################
###########################################################################################################
import os, time, sys

sys.stdout.flush()
print('VERIFICANDO SISTEMA...\n')
time.sleep(1)
os.system('pip install paramiko')
os.system('pip install openpyxl')
os.system('pip install termcolor')

###########################################################################################################
################################# IMPORTANDO DEMAIS PARAMETROS NECESSÁRIOS ################################
###########################################################################################################
import json, threading, re, shutil, tarfile, subprocess, platform, paramiko, socket
from paramiko import SSHClient, AutoAddPolicy, transport, SecurityOptions, ssh_exception
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from termcolor import colored
from getpass import getpass

###########################################################################################################
############### DETECTANDO SISTEMA OPERACIONAL PARA APARENCIA INICIAL E INFORMAÇÃO AO USUARIO #############
###########################################################################################################
print("\n\n\n------> DETECTANDO SISTEMA OPERACIONAL <------")
if platform.system() == "Windows":
    print("Sistema Operacional", end=' ')
    print(colored("Windows", 'blue'))
    time.sleep(2)
    os.system('cls')

elif platform.system() == "Linux":
    print("Sistema Operacional", end=' ')
    print(colored("Linux", 'green'))
    time.sleep(2)
    os.system('clear')
    
else:
    print(colored("SISTEMA OPERACIONAL INCOMPATIVEL COM PROGRAMA DE RESGATE!!!", 'red'))
    print('FECHANDO...')
    time.sleep(5)
    exit()
    
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

###########################################################################################################
################################### VERIFICA SE EXISTE OU CRIA AS PASTAS ##################################
###########################################################################################################

############## VERIFICANDO SE UNIDADE D: EXISTE ##############
if platform.system() == "Windows":
    if os.path.isdir("G:"):
        g_C_D = 'G:'
    else:
        g_C_D = 'C:'

############## VERIFICANDO PASTA PADRAO LINUX ##############
elif platform.system() == "Linux":
    USUARIO = os.getlogin()
    HOME = f'/media/{USUARIO}/PRODATA'

    if os.path.isdir(HOME):
        g_C_D = f'/media/{USUARIO}/PRODATA'
    
    else:
        FDISK = subprocess.run(["fdisk -l"], shell=True, capture_output=True, text=True)
        SDB = FDISK.stdout 
        SDB2 = SDB.find("/dev/sdb2")
        
        if SDB2 >= 0:
            os.system('umount /dev/sdb2')
            os.system(f'mkdir /media/{USUARIO}/PRODATA')
            os.system(f'mount /dev/sdb2 /media/{USUARIO}/PRODATA')
            g_C_D = f'/media/{USUARIO}/PRODATA'

        elif os.path.isdir(f"/home/{USUARIO}/Documentos"):
            g_C_D = f'/home/{USUARIO}/Documentos'

        elif os.path.isdir(f"/home/{USUARIO}/Documents"):
            g_C_D = f'/home/{USUARIO}/Documents'

        else:
            print(colored('---> FALHA AO DEFINIR PASTA PADRÃO!!! <---\n\n', 'red'))
            print(colored("---> OBS: CRIE UMA PASTA COM O NOME DE 'PRODATA' EM /media/'seu_usuario'/\n", 'white'))
            print(colored("          EXEMPLO: /media/'seu_usuario'/PRODATA!!! <---", 'yellow'))
            print(colored("          E REINICIE O PROGRAMA.", 'white'))
            print(colored("\n\nREINICIAR?", 'yellow'))
            REINICAR2 = input('Digite S para SIM ou ENTER para Fechar...\n').upper()
            if REINICAR2 == 'S':
                print('REINICIANDO...')
                os.system('python3 RESGATE.py')
            elif REINICAR2 == 'SIM':
                print('REINICIANDO...')
                os.system('python3 RESGATE.py')
            else:
                print('FECHANDO...\n')
                time.sleep(2)
                os.system('pkill terminal')

###########################################################################################################
################################################## GLOBALS ################################################
###########################################################################################################
lastValidHost : str
validPort : int
connections = []
ssh_sessions : list

g_ips : dict
g_ports : dict
g_users : dict
g_passwords : dict
g_passphrase : str

###########################################################################################################
####################################### DETECTA OS PRA INICIAR RESGATE ####################################
###########################################################################################################
def detect_os():
    if platform.system() == "Windows":
        windows_processar(get_ssh_session())
    elif platform.system() == "Linux": 
        linux_processar(get_ssh_session())

###########################################################################################################
########################################## CARREGA AS CONF DO JSON ########################################
###########################################################################################################
def carregar_configs():
    global g_ips, g_users, g_passwords, g_ports, g_passphrase

    ########## CONFIGURANDO ARQUIVO JSON ##########
    configFile = os.path.realpath(__file__).replace("py", "json")
    with open(configFile) as file:
        data = json.load(file)

    g_ips = data['hosts']
    g_users = data['usuario']
    g_passwords = data['senhas']
    g_ports = data['portas']
    g_passphrase = data['passphrase']

###########################################################################################################
############################################# CONECTA PELA SENHA ##########################################
###########################################################################################################
def thread_connect_password(pwdNum):
    global lastValidHost, validPort, ssh_sessions, g_users, g_passwords

    client = SSHClient()
    client.load_system_host_keys()
    client.set_missing_host_key_policy(AutoAddPolicy())
    
    try:
        client.connect(lastValidHost, username=g_users, password=g_passwords[str(pwdNum)], timeout=2.0)
        ssh_sessions.append(client)
    except:
        client.close()
        
###########################################################################################################
############################################## CONECTA PELA KEY ###########################################
###########################################################################################################
def ssh_connect_key(disabled_alg=None) -> bool:
    global lastValidHost, validPort, ssh_sessions, g_users, g_SSH, g_passphrase

    client = SSHClient()
    client.load_system_host_keys()
    client.set_missing_host_key_policy(AutoAddPolicy())
    g_SSH = paramiko.RSAKey.from_private_key_file('C:\\RESGATE\\PRIVATEKEY', password=g_passphrase)

    try:
        client.connect(lastValidHost, username=g_users, pkey=g_SSH, port=validPort, timeout=2.0, disabled_algorithms=disabled_alg)
        ssh_sessions.append(client)
        return True
    except:
        client.close()
    
    return False

###########################################################################################################
########################################## OBTENDO IP DO VALIDADOR ########################################
###########################################################################################################
def thread_get_ip(ip, port):
    global lastValidHost, validPort

    try:
        s = socket.create_connection(address=(ip, port), timeout=2.0)
        s.close()
        lastValidHost = ip
        validPort = port
    except:
        pass

###########################################################################################################
############################################# CONECTA PELA SENHA ##########################################
###########################################################################################################
def get_host_ip_address(ports=[]):
    global lastValidHost, validPort, g_ips
    
    threads = []
    validPort = -1
    lastValidHost = ""
    
    index = 0
    for port in ports:
        for _, ip in g_ips.items():
            t = threading.Thread(target=thread_get_ip, args=(ip, port))
            threads.append((index, t))
            t.start()
            index = index + 1
    
    for index, thread in threads:
        try:
            thread.join()
        except:
            pass

###########################################################################################################
######################################## COMANDOS PARA SISTEMA OPERACIONAL ################################
###########################################################################################################
def system_call(command):
    try:
        output = subprocess.check_output(command, stderr=subprocess.STDOUT, universal_newlines=True)
    except subprocess.CalledProcessError as e:
        output = e.output.decode()
    return output

###########################################################################################################
################################### VERIFICA SE VALIDADOR ESTA NA MESMA REDE ##############################
###########################################################################################################
def windows_is_local_networked_host(ip) -> bool:
    try:
        output = system_call(f"ping -4 -n 1 -i 1 -w 50 {ip}")
        ###TTL SETA SE FOR 1. SE O VALIDADOR ESTIVER EM OUTRA REDE O ICMP/ping SERA -1.
    except Exception as e:
        print(str(e))
        return False

    for line in output.split('\r'):
        if(line.lower().rfind("ttl=") != -1):
            return True
    
    return False

###########################################################################################################
############################################# INICIA SESSAO SSH ###########################################
###########################################################################################################
def get_ssh_session() -> SSHClient:
    global ssh_sessions, g_passwords
    ssh_sessions = []
    
    if(ssh_connect_key()):
        #SSH COM ALGORITIMOS
        pass
    elif(ssh_connect_key({'pubkeys': ['rsa-sha2-256', 'rsa-sha2-512']})):
        #SSH SEM ALGORITIMO
        pass
    elif(ssh_connect_key({'pubkeys': ['rsa-sha2-256']})):
        #SSH SEM ALGORITIMO
        pass
    elif(ssh_connect_key({'pubkeys': ['rsa-sha2-512']})):
        #SSH SEM ALGORITIMO
        pass
    else:
        #SSH COM TODAS AS SENHAS
        threads = []
        index = 0
        for p in range(len(g_passwords)):
            t = threading.Thread(target=thread_connect_password, args=(p,))
            threads.append(t)
            t.start()
            index = index + 1
        
        for t in threads:
            t.join()
        
        del threads
    
    if(len(ssh_sessions) > 0):
        session = ssh_sessions[0]
        del ssh_sessions
        return session
    else:
        conexao_falhou()
    
    return None

###########################################################################################################
############################################## COMANDOS EM SSH ############################################
###########################################################################################################
def ssh_command(client:SSHClient, cmd:str):
    stdin, stdout, stderr = client.exec_command(cmd)
    
    if(stdout):
        print(f"{stdout.read().decode('utf8')}", end='')
    if(stderr):
        print(f"{stderr.read().decode('utf8')}", end='')

    stdin.close()
    
###########################################################################################################
############################################## FALHA NA CONEXAO ###########################################
###########################################################################################################
def conexao_falhou():
    global lastValidHost, g_C_D, g_SP_CMT2, g_EMPRESA, g_OS, g_NS

    ########### SE FALHAR QUALQUER TENTATIVA DE CONEXAO NO WINDOWS ###########
    print(colored('FALHA NA COMUNICAÇÃO:', "red"))
    print(colored('NÃO FOI POSSÍVEL ESTABELECER COMUNICAÇÃO COM VALIDADOR', "yellow"))#", end="")
    #print(f'{lastValidHost}')
    
    if platform.system() == "Windows":
        if os.path.exists(g_C_D+f'\\BKP\\RESGATE-{ANO}.xlsx'):
            pass
        else:
            # CRIANDO NOVA PLANILHA
            RESGATE = Workbook()
            INTERVENCAO = RESGATE.active
            INTERVENCAO.title = 'INTERVENCAO'
            RESGATE.save(g_C_D+f'\\BKP\\RESGATE-{ANO}.xlsx')
            # ESCREVENDO TITULO
            RESGATE = load_workbook(g_C_D+f'\\BKP\\RESGATE-{ANO}.xlsx')
            INTERVENCAO = RESGATE['INTERVENCAO']
            TITULO = ['OK', 'EMPRESA', 'O. DE SERV.', 'SERIAL', 'TEM ARQUIVOS', 'SEQUENCIAL DOS ARQUIVOS', 'DATA', 'OBSERVAÇÃO', 'DEFEITO CONSTATADO']
            INTERVENCAO.append(TITULO)
            #ALINHAMENTO
            ALINHAMENTO = Alignment(horizontal='center')
            CABECALHO_TOTAL = ['A1', 'B1', 'C1', 'D1', 'E1', 'F1', 'G1', 'H1', 'I1']
            for cell in CABECALHO_TOTAL:
                INTERVENCAO[cell].alignment = ALINHAMENTO
            # CABEÇALHO EM NEGRITO, FUNDO AZUL E BORDAS
            COR_FUNDO = PatternFill(start_color='5399FF', end_color='5399FF', fill_type='solid')
            NEGRITO = Font(bold=True)
            BORDA = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            CABECALHO = ['B1', 'C1', 'D1', 'E1', 'F1', 'G1', 'H1', 'I1']
            for cell in CABECALHO:
                INTERVENCAO[cell].fill = COR_FUNDO
                INTERVENCAO[cell].font = NEGRITO
                INTERVENCAO[cell].alignment = ALINHAMENTO
                INTERVENCAO[cell].border = BORDA
            # SALVANDO PLANILHA RESGATE DO ANO
            RESGATE.save(g_C_D+f'\\BKP\\RESGATE-{ANO}.xlsx')
            print (f'PLANILHA PRA RESGATE-{ANO} CRIADA COM SUESSO!')
            
        ############## VERIFICA SE EXISTE OU CRIA PASTA PARA SPTRANS OU CMT-PROJETOS OU TOP-BOM ##############
        if os.path.isdir(g_C_D+'\\BKP\\'+g_SP_CMT2):
            time.sleep(1)
        else:
            os.mkdir(g_C_D+'\\BKP\\'+g_SP_CMT2)
            time.sleep(1)

        ############## VERIFICA SE EXISTE OU CRIA PASTA PARA EMPRESA ##############
        if os.path.isdir(g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA):
            time.sleep(1)
        else:
            os.mkdir(g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA)
            time.sleep(1)

        ############## VERIFICA SE EXISTE OU CRIA PASTA PARA OS ##############
        if os.path.isdir(g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA):
            time.sleep(1)
        else:
            os.mkdir(g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA)
            time.sleep(1)
        
        ### GRAVANDO INFORMAÇÕES NO TXT DE LOCALIZAÇÃO ###
        LOCALIZADOR = open(g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\'+g_OS+' - '+g_NS+'.txt', 'a')
        LOCALIZADOR.write('\n-> FALHA NA COMUNICAÇÃO COM O VALIDADOR')
        LOCALIZADOR.close()
        
        ### CARREGANDO PLANILHA DE RESGATE ###
        RESGATE = load_workbook(g_C_D+f'\\BKP\\RESGATE-{ANO}.xlsx')
        PLANILHA = RESGATE['INTERVENCAO']
        try:
            ### ADCIONINANDO INFORMACOES NA PLANILHA ###
            RESGATE_INFO = ['F', g_EMPRESA, g_OS, g_NS, 'NAO', ' ', DATAXLSX, 'FALHA NA COMUNICAÇÃO COM O VALIDADOR']
            PLANILHA.append(RESGATE_INFO)
            NLINHA = PLANILHA.max_row
            ULTIMA_LINHA = [f'A{NLINHA}', f'B{NLINHA}', f'C{NLINHA}', f'D{NLINHA}', f'E{NLINHA}', f'F{NLINHA}', f'G{NLINHA}', f'H{NLINHA}', f'I{NLINHA}']
            ULTIMA_LINHA2 = [f'B{NLINHA}', f'C{NLINHA}', f'D{NLINHA}', f'E{NLINHA}', f'F{NLINHA}', f'G{NLINHA}', f'H{NLINHA}', f'I{NLINHA}']
            #BORDA = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            for cell in ULTIMA_LINHA2:
                PLANILHA[cell].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            for cell in ULTIMA_LINHA:
                PLANILHA[cell].alignment = Alignment(horizontal='center')
                
            RESGATE.save(g_C_D+f'\\BKP\\RESGATE-{ANO}.xlsx')
            print(colored('\n   GRAVADA INFORMAÇÕES NA PLANILHA!!!', 'green'))
        except:
            print(colored('\nFALHA AO GRAVAR INFORMAÇÕES NA PLANILHA!!!', 'red'))
            print(colored('VERIFIQUE SE A PLANILHA ESTA ABERTA E FECHEA.', 'yellow'))
            pass
                
        ### VERIFICAR SE REINICIA OU NAO ###
        print(colored("\n\nREINICIAR?", 'yellow'))
        REINICAR_ERRO = input('Digite S para SIM ou ENTER para Fechar...\n').upper().strip()
        if REINICAR_ERRO == 'S':
            print('REINICIANDO...')
            os.system('python resgate.py')
        elif REINICAR_ERRO == 'SIM':
            print('REINICIANDO...')
            os.system('python resgate.py')
        elif REINICAR_ERRO == 'SS':
            print('REINICIANDO...')
            os.system('python resgate.py')
        else:
            print('FECHANDO...\n')
            os.system('TASKKILL /IM python.exe /T /F')

    elif platform.system() == "Linux":
        if os.path.exists(g_C_D+f'\\BKP\\RESGATE-{ANO}.xlsx'):
            pass
        else:
            # CRIANDO NOVA PLANILHA
            RESGATE = Workbook()
            INTERVENCAO = RESGATE.active
            INTERVENCAO.title = 'INTERVENCAO'
            RESGATE.save(g_C_D+f'\\BKP\\RESGATE-{ANO}.xlsx')
            # ESCREVENDO TITULO
            RESGATE = load_workbook(g_C_D+f'\\BKP\\RESGATE-{ANO}.xlsx')
            INTERVENCAO = RESGATE['INTERVENCAO']
            TITULO = ['OK', 'EMPRESA', 'O. DE SERV.', 'SERIAL', 'TEM ARQUIVOS', 'SEQUENCIAL DOS ARQUIVOS', 'DATA', 'OBSERVAÇÃO', 'DEFEITO CONSTATADO']
            INTERVENCAO.append(TITULO)
            #ALINHAMENTO
            ALINHAMENTO = Alignment(horizontal='center')
            CABECALHO_TOTAL = ['A1', 'B1', 'C1', 'D1', 'E1', 'F1', 'G1', 'H1', 'I1']
            for cell in CABECALHO_TOTAL:
                INTERVENCAO[cell].alignment = ALINHAMENTO
            # CABEÇALHO EM NEGRITO, FUNDO AZUL E BORDAS
            COR_FUNDO = PatternFill(start_color='5399FF', end_color='5399FF', fill_type='solid')
            NEGRITO = Font(bold=True)
            BORDA = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            CABECALHO = ['B1', 'C1', 'D1', 'E1', 'F1', 'G1', 'H1', 'I1']
            for cell in CABECALHO:
                INTERVENCAO[cell].fill = COR_FUNDO
                INTERVENCAO[cell].font = NEGRITO
                INTERVENCAO[cell].alignment = ALINHAMENTO
                INTERVENCAO[cell].border = BORDA
            # SALVANDO PLANILHA RESGATE DO ANO
            RESGATE.save(g_C_D+f'\\BKP\\RESGATE-{ANO}.xlsx')
            print (f'PLANILHA PRA RESGATE-{ANO} CRIADA COM SUESSO!')
            
        ### GRAVANDO INFORMAÇÕES NO TXT DE LOCALIZAÇÃO ###
        LOCALIZADOR = open(g_C_D+'/BKP/'+g_SP_CMT2+'/'+g_EMPRESA+'/'+g_OS+DATA+'/'+g_OS+' - '+g_NS+'.txt', 'a')
        LOCALIZADOR.write('\n-> FALHA NA COMUNICAÇÃO COM O VALIDADOR')
        LOCALIZADOR.close()
               
        ### CARREGANDO PLANILHA DE RESGATE ###
        RESGATE = load_workbook(g_C_D+f'/BKP/RESGATE-{ANO}.xlsx')
        PLANILHA = RESGATE['INTERVENCAO']
        try:
            ### ADCIONINANDO INFORMACOES NA PLANILHA ###
            RESGATE_INFO = ['F', g_EMPRESA, g_OS, g_NS, 'NAO', ' ', DATAXLSX, 'FALHA NA COMUNICAÇÃO COM O VALIDADOR']
            PLANILHA.append(RESGATE_INFO)
            
            NLINHA = PLANILHA.max_row
            ULTIMA_LINHA = [f'A{NLINHA}', f'B{NLINHA}', f'C{NLINHA}', f'D{NLINHA}', f'E{NLINHA}', f'F{NLINHA}', f'G{NLINHA}', f'H{NLINHA}', f'I{NLINHA}']
            ULTIMA_LINHA2 = [f'B{NLINHA}', f'C{NLINHA}', f'D{NLINHA}', f'E{NLINHA}', f'F{NLINHA}', f'G{NLINHA}', f'H{NLINHA}', f'I{NLINHA}']
            #BORDA = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            for cell in ULTIMA_LINHA2:
                PLANILHA[cell].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            for cell in ULTIMA_LINHA:
                PLANILHA[cell].alignment = Alignment(horizontal='center')
            RESGATE.save(g_C_D+f'/BKP/RESGATE-{ANO}.xlsx')
            print(colored('\n   GRAVADA INFORMAÇÕES NA PLANILHA!!!', 'green'))
        except:
            print(colored('\nFALHA AO GRAVAR INFORMAÇÕES NA PLANILHA!!!', 'red'))
            print(colored('VERIFIQUE SE A PLANILHA ESTA ABERTA E FECHEA.', 'yellow'))
            pass
        
        ### VERIFICAR SE REINICIA OU NAO ###
        print(colored("\n\nREINICIAR?", 'yellow'))
        REINICAR_ERRO = input('Digite S para SIM ou ENTER para Fechar...\n').upper().strip()
        if REINICAR_ERRO == 'S':
            print('REINICIANDO...')
            os.system('python3 RESGATE.py')
        elif REINICAR_ERRO == 'SIM':
            print('REINICIANDO...')
            os.system('python3 RESGATE.py')
        elif REINICAR_ERRO == 'SS':
            print('REINICIANDO...')
            os.system('python3 RESGATE.py')
        else:
            print('FECHANDO...\n')
            time.sleep(2)
            os.system('pkill terminal')

###########################################################################################################
################################################# WINDOWS #################################################
###########################################################################################################
def windows_processar(client : SSHClient):
    global g_EMPRESA, g_OS, g_NS, g_C_D, g_SP_CMT2

    ########################################### CONSTANTES PARA PASTAS DO WINDOWS ###########################################
    FILE_UD = ('/opt/v750/system/ud/Ud.tar.gz')
    FILE_SD = ('/media/mmcblk0p1/opt/v750/system/ud/System.tar.gz')
    FILE_PCGAR = ('/opt/v750/system/ud/PCGAR.tar.gz')
    FILE_PCGAR2 = ('/media/mmcblk0p1/opt/v750/system/ud/PCGAR.tar.gz')
    FILE_TG = ('/opt/v750/system/ud/TG.tar.gz')
    FILE_TG2 = ('/media/mmcblk0p1/opt/v750/system/ud/TG.tar.gz')
    FILE_TOP = ('/opt/v750/system/ud/TOP.tar.gz')
    FILE_BOM = ('/001/opt/v750/system/ud/BOM.tar.gz')
    FILE_TG_BOM = ('/001/opt/v750/system/ud/TG-BOM.tar.gz')
    FILE_TG_BOM2 = ('/media/mmcblk0p1/001/opt/v750/system/ud/TG-BOM.tar.gz')
    FILE_SD_TOP = ('/media/mmcblk0p1/mmcblk0p1.tar.gz')

    PASTA_INI = (g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\')
    LOCAL_UD = (g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\Ud.tar.gz')
    LOCAL_SD = (g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\System.tar.gz')
    LOCAL_PCGAR = (g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\PCGAR.tar.gz')
    LOCAL_TG = (g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\TG.tar.gz')
    LOCAL_TOP = (g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\TOP.tar.gz')
    LOCAL_BOM = (g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\BOM.tar.gz')
    LOCAL_TG_BOM = (g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\TG-BOM.tar.gz')
    LOCAL_SD_TOP = (g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\mmcblk0p1.tar.gz')
    
    DESC_UD = (g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\Ud\\')
    DESC_SD = (g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\System\\')
    DESC_TOP = (g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\TOP\\')
    DESC_TG_BOM = (g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\TG-BOM\\')
    DESC_BOM = (g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\BOM\\')
    DESC_SD_TOP = (g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\mmcblk0p1\\')

    TXT = (g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\serial.txt')
    TXT2 = (g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\serialread.txt')
    FILE_TXT = ('/root/serial.txt')
    FILE_TXT_CMT = ('/root/serialread.txt')
    PASTA_TXT = (g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\serial.txt')
    PASTA_TXT_CMT = (g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\serialread.txt')

    ############################## CONEXAO ESTABELECIDA ##############################
    print(colored(f"\n   ---> CONECTADO NO IP {lastValidHost} <--- \n", 'light_cyan', attrs=['bold']))    
        
    ############## VERIFICA SE EXISTE OU CRIA PLANILHA PARA RELATORIO DE BKP ##############
    if os.path.exists(g_C_D+f'\\BKP\\RESGATE-{ANO}.xlsx'):
        pass
    else:
        # CRIANDO NOVA PLANILHA
        RESGATE = Workbook()
        INTERVENCAO = RESGATE.active
        INTERVENCAO.title = 'INTERVENCAO'
        RESGATE.save(g_C_D+f'\\BKP\\RESGATE-{ANO}.xlsx')
        # ESCREVENDO TITULO
        RESGATE = load_workbook(g_C_D+f'\\BKP\\RESGATE-{ANO}.xlsx')
        INTERVENCAO = RESGATE['INTERVENCAO']
        TITULO = ['OK', 'EMPRESA', 'O. DE SERV.', 'SERIAL', 'TEM ARQUIVOS', 'SEQUENCIAL DOS ARQUIVOS', 'DATA', 'OBSERVAÇÃO', 'DEFEITO CONSTATADO']
        INTERVENCAO.append(TITULO)
        #ALINHAMENTO
        ALINHAMENTO = Alignment(horizontal='center')
        CABECALHO_TOTAL = ['A1', 'B1', 'C1', 'D1', 'E1', 'F1', 'G1', 'H1', 'I1']
        for cell in CABECALHO_TOTAL:
            INTERVENCAO[cell].alignment = ALINHAMENTO
        # CABEÇALHO EM NEGRITO, FUNDO AZUL E BORDAS
        COR_FUNDO = PatternFill(start_color='5399FF', end_color='5399FF', fill_type='solid')
        NEGRITO = Font(bold=True)
        BORDA = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        CABECALHO = ['B1', 'C1', 'D1', 'E1', 'F1', 'G1', 'H1', 'I1']
        for cell in CABECALHO:
            INTERVENCAO[cell].fill = COR_FUNDO
            INTERVENCAO[cell].font = NEGRITO
            INTERVENCAO[cell].alignment = ALINHAMENTO
            INTERVENCAO[cell].border = BORDA
        # SALVANDO PLANILHA RESGATE DO ANO
        RESGATE.save(g_C_D+f'\\BKP\\RESGATE-{ANO}.xlsx')
        print (f'RESGATE-{ANO} criada com sucesso!\n')

    ############## VERIFICA SE EXISTE OU CRIA PASTA PARA SPTRANS OU CMT-PROJETOS OU TOP-BOM ##############
    if os.path.isdir(g_C_D+'\\BKP\\'+g_SP_CMT2):
        time.sleep(1)
    else:
        os.mkdir(g_C_D+'\\BKP\\'+g_SP_CMT2)
        print ('Pasta para '+ g_SP_CMT2+' criada com sucesso!')
        print('')
        time.sleep(1)

    ############## VERIFICA SE EXISTE OU CRIA PASTA PARA EMPRESA ##############
    if os.path.isdir(g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA):
        print ('Já existe uma pasta de BKP para a empresa', g_EMPRESA)
        print('')
        time.sleep(1)
    else:
        os.mkdir(g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA)
        print ('Pasta para a empresa', g_EMPRESA, 'criada com sucesso!')
        print('')
        time.sleep(1)

    ############## VERIFICA SE EXISTE OU CRIA PASTA PARA OS ##############
    if os.path.isdir(g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA):
        print ('Já existe uma pasta para a OS', g_OS)
        print('')
        time.sleep(1)
    else:
        os.mkdir(g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA)
        print ('Pasta para a OS', g_OS, 'criada com sucesso!')
        print('')
        time.sleep(1)

    ############## VERIFICA SE EXISTE OU CRIA TXT PARA LOCALIZAR PASTA OS ##############
    LOCALIZADOR = open(g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\'+g_OS+' - '+g_NS+'.txt', 'a')
    #INSERINDO INFORMAÇÕES NO TXT
    LOCALIZADOR = open(g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\'+g_OS+' - '+g_NS+'.txt', 'w')
    LOCALIZADOR.write ('EMPRESA: '+g_EMPRESA +'\n' +'OS: '+g_OS +'\n' + 'NS: '+g_NS+'\n' +'DATA: '+DATAXLSX)
    LOCALIZADOR.close()
    print(colored('TXT para localizar OS criada com sucesso!', 'green'))
    print('')
    time.sleep(1)
    
    ############################## INVOCANDO XTERM PRA ENVIAR COMANDOS ##############################
    client.invoke_shell() 
    
    ############################## VERIFICAR EXECUTAVEL ##############################
    stdin, stdout, stderr = client.exec_command('cd /root/app\n''ls Mercury-v3680hf')
    time.sleep(0.5)
    EXECUTAVELRTC1 = stdout.read().decode('utf8')

    stdin, stdout, stderr = client.exec_command('cd /root/app\n''ls Control-v3695')
    time.sleep(0.5)
    EXECUTAVELRTC2 = stdout.read().decode('utf8')

    stdin, stdout, stderr = client.exec_command('cd /root/app\n''ls Mercury-v3695')
    time.sleep(0.5)
    EXECUTAVELRTC3 = stdout.read().decode('utf8')

    stdin, stdout, stderr = client.exec_command('cd /root/app\n''ls Mercury-v3690hf')
    time.sleep(0.5)
    EXECUTAVELRTC4 = stdout.read().decode('utf8')

    stdin, stdout, stderr = client.exec_command('cd /root/app\n''ls Mercury-v3680')
    time.sleep(0.5)
    EXECUTAVELRTC5 = stdout.read().decode('utf8')

    stdin, stdout, stderr = client.exec_command('cd /root/app\n''ls Mercury-v3690')
    time.sleep(0.5)
    EXECUTAVELRTC6 = stdout.read().decode('utf8')

    if EXECUTAVELRTC1[:-1] == "Mercury-v3680hf":
            EXECUTAVELRTC = "Mercury-v3680hf"
        
    elif EXECUTAVELRTC2[:-1] == "Control-v3695":
            EXECUTAVELRTC = "Control-v3695"
        
    elif EXECUTAVELRTC3[:-1] == "Mercury-v3695":
            EXECUTAVELRTC = "Mercury-v3695"

    elif EXECUTAVELRTC4[:-1] == "Mercury-v3690hf":
            EXECUTAVELRTC = "Mercury-v3690hf"
        
    elif EXECUTAVELRTC5[:-1] == "Mercury-v3680":
            EXECUTAVELRTC = "Mercury-v3680"
        
    elif EXECUTAVELRTC6[:-1] == "Mercury-v3690":
            EXECUTAVELRTC = "Mercury-v3690"
    
    ############################## CONFIGURANDO RTC PARA ELIMINAR ERRO ##############################
    if g_SP_CMT2 == 'CMT-PROJETOS':
        stdin, stdout, stderr = client.exec_command(f'pkill -f {EXECUTAVELRTC}')
        stdin, stdout, stderr = client.exec_command('pkill -f Connection-v3680hf')
        stdin, stdout, stderr = client.exec_command('killall run_app run_connection run_qrdaemon run_abt')
        stdin, stdout, stderr = client.exec_command('killall run start')
    else:
        pass
    
    print(colored("---> CONFIGURANDO RTC <---", 'blue', attrs=['bold']))
    ssh_command(client, f"date {RTC}")
    time.sleep(1)
    stdin, stdout, stderr = client.exec_command("hwclock -w")
    stdin, stdout, stderr = client.exec_command("hwclock -w -f /dev/rtc")
    stdin, stdout, stderr = client.exec_command("hwclock -w -f /dev/rtc1")
    time.sleep(1)
    print(colored('!!! RTC - OK !!!\n', 'green',))
    
    ############################## VERIFICAR SE É EQUIPAMENTO AUXILIAR DE BKP ##############################
    # VALIDADOR DEVE CONTEM O SEGUINTE ARQUIVO:
    # VALIDADOR V3680 = '####### VALIDADOR RESGATE SDCARD #######'
    # VALIDAODR V369X = '##### COMPULAB DE BKP #####'
    
    stdin, stdout, stderr = client.exec_command('cd /\n''ls')
    BKP = stdout.read().decode('utf8')
    BK = BKP.find("VALIDADOR RESGATE SDCARD")
    INFO_PCGAR_BOM = ' '
    SEQUENCIAL_INFO = ' '
    SEQUENCIAL_INFO2 = ' '
    SEQUENCIAL_PCGAR = ' '
    INFO_PCGAR_TG = ' '
    OBS_PLAN = ' '
    SD_PLAN = ' '    
    
    if BK >= 0:
            print(colored('!!! EQUIPAMENTO AUXILIAR DE BKP !!!', 'yellow', attrs=['bold']))
            print(colored('Serão resgatadas as informações somente do SDCard se possível!\n', attrs=['bold']))
            time.sleep(2)
            BKP_PLAN = 'VALIDADOR DE RESGATE'
            ### VERIFICANDO A EXISTENCIA DO SDCARD ###
            print(colored("---> VERIFICANDO SDCARD <---", 'blue', attrs=['bold']))
            time.sleep(1)
            stdin, stdout, stderr = client.exec_command("fdisk -l /dev/mmcblk0p1")
            SDCARD = stdout.read().decode('utf8')
            SD = SDCARD.find('/dev/mmcblk0p1')
            print(SDCARD[1:29], end='')
            
            if SD == -1:
                ### SDCARD INOPERANTE OU NAO EXISTE ###
                print(colored('\r   !!! SDCard não detectado !!!\n\n', 'red', attrs=['bold']))
                time.sleep(1)
                SD_PLAN = 'SDCARD COM DEFEITO'
                INFO_PCGAR_TG = 'NAO'

                LOCALIZADOR = open(g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\'+g_OS+' - '+g_NS+'.txt', 'a')
                LOCALIZADOR.write('\n-> VALIDADOR COM DEFEITO OU BLOQUEADO - UTILIZADO VALIDADOR AUXILIAR DE BKP (SOMENTE SDCARD)\n-> SDCARD COM DEFEITO')
                LOCALIZADOR.close()
                print()

            else:
                print(colored('\n    !!! SDCard - OK !!!\n\n', 'green', attrs=['bold']))
                print(colored(f'---> VARIFICANDO SE VALIDADOR DE BKP É TG OU PCGAR <---', 'blue', attrs=['bold']))
                stdin, stdout, stderr = client.exec_command('cd /root/app/\n''ls SPTrans-v3068')
                time.sleep(1)
                EXECUTAVELY = stdout.read().decode('utf8')
                
                stdin, stdout, stderr = client.exec_command('cd /root/app/\n''ls SPTrans-v3680hf')
                time.sleep(1)
                EXECUTAVELYY = stdout.read().decode('utf8')
                
                stdin, stdout, stderr = client.exec_command('cd /root/app/\n''ls Mercury-v3695')
                time.sleep(1)
                EXECUTAVELZ = stdout.read().decode('utf8')
                OBS_PLAN = 'RESGATADO SOMENTE DO SDCARD'

                LOCALIZADOR = open(g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\'+g_OS+' - '+g_NS+'.txt', 'a')
                LOCALIZADOR.write("\n-> VALIDADOR COM DEFEITO OU BLOQUEADO - UTILIZADO VALIDADOR AUXILIAR DE BKP (SOMENTE SDCARD)")
                LOCALIZADOR.close()

                if EXECUTAVELY[:-1] == "SPTrans-v3068":
                    print(colored('VALIDADOR DE BKP: SPTRANS\n', attrs=['bold']))
                    EXECUTAVEL_A = "SPTrans-v3068"
                    time.sleep(2)
                    print(colored('---> INICIANDO BKP DA PASTA SYSTEM EM SDCARD <---', 'blue', attrs=['bold']))
                    stdin, stdout, stderr = client.exec_command('find /media/mmcblk0p1/opt/v750/system/ud')
                    time.sleep(2)
                    try:
                        VERFSD = stdout.read().decode('utf8')
                    except:
                        VERFSD = 'SDCARD COM DEFEITO'
                    
                    if VERFSD[0:35] == "/media/mmcblk0p1/opt/v750/system/ud":
                        stdin, stdout, stderr = client.exec_command('cd /media/mmcblk0p1/opt/v750/system/ud/backup/ \n''ls')
                        VERFSD2 = stdout.read().decode('utf8')
                        if VERFSD2 == "":
                            print(colored('AVISO: Pasta BACKUP do SDCARD está vazia!', 'yellow', attrs=['bold']))
                        else:
                            pass

                        stdin, stdout, stderr = client.exec_command("cd /media/mmcblk0p1/opt/v750/system/ud \n""tar --exclude='photos' -cf System.tar.gz *")
                        time.sleep(10)
                        print(colored('     Arquivo TAR do SDCARD:', attrs=['bold']))
                        ssh_command(client, "cd /media/mmcblk0p1/opt/v750/system/ud \n""ls -lh System*")
                        print('DOWNLOAD: ...', end=' ')
                        time.sleep(2)
                        ### DOWNLOAD DO ARQUIVO ###
                        ftp = client.open_sftp()
                        ftp.get(FILE_SD, LOCAL_SD)
                        ftp.close()
                        print(colored('\rDOWNLOAD: OK  ', 'green', attrs=['bold']))
                        print()

                        print(colored('---> VERIFICANDO ARQUIVOS PCGAR DE BKP <---', 'blue', attrs=['bold']))
                        #ssh_command(client,"cd /opt/v750/system/ud/\n""ls File_002_*\n")
                        stdin, stdout, stderr = client.exec_command("cd /media/mmcblk0p1/opt/v750/system/ud/\n""du -hsb File_112_00000\n")
                        time.sleep(1)
                        KB_00000 = stdout.read().decode('utf8')
                        KB = KB_00000.find('File_112_00000')

                        if KB == -1:
                            print(colored("SEM ARQUIVO DE BKP PARA RESGATAR!!!\n", 'red', attrs=['bold']))
                            time.sleep(2)
                            INFO_PCGAR_TG = 'NAO'
                            
                        elif KB >= 0:
                            if (int(KB_00000[:-15])) <= 1000:
                                print(KB_00000[:-15]+'File_112_00000') 
                                print(colored(' --> File_112_00000 MENOR QUE 1000Kb <--\n', 'red', attrs=['bold']))
                                time.sleep(1.5)
                            else:
                                print(colored(' --> INICIANDO RESGATE DO ARQUIVO File_112_00000 <--', 'green', attrs=['bold']))
                                print(KB_00000[:-15]+'File_112_00000')
                                print(colored('\n -> VERIFICANDO SEQUENCIAL <-',attrs=['bold']))
                                stdin, stdout, stderr = client.exec_command("cd /media/mmcblk0p1/opt/v750/system/ud/backup\n""ls -t File_112_*")
                                time.sleep(2)
                                BACKUP = stdout.read().decode('utf8')
                                SEQUENCIALBKP = (BACKUP[9:14])

                                try:
                                    SEQUENCIALBKPINT = (int(SEQUENCIALBKP)+1)
                                except:
                                    SEQUENCIALBKPINT = (int(g_NS))

                                if (len(str(SEQUENCIALBKPINT))) == 5:
                                    pass
                                elif (len(str(SEQUENCIALBKPINT))) == 4:
                                    SEQUENCIALBKPINT = '0'+(str(SEQUENCIALBKPINT))
                                elif (len(str(SEQUENCIALBKPINT))) == 3:
                                    SEQUENCIALBKPINT = '00'+(str(SEQUENCIALBKPINT))
                                elif (len(str(SEQUENCIALBKPINT))) == 2:
                                    SEQUENCIALBKPINT = '000'+(str(SEQUENCIALBKPINT))
                                elif (len(str(SEQUENCIALBKPINT))) == 1:
                                    SEQUENCIALBKPINT = '0000'+(str(SEQUENCIALBKPINT))

                                print(colored(f'SEQUENCIAL:', attrs=['bold']), end=' ')
                                print(colored(f'{SEQUENCIALBKPINT}', 'green'))
                                time.sleep(1)

                                ### RENOMEANDO FILE_112_00000 ###
                                print(colored("\n    --> RENOMEANDO FILE 112 <--", 'blue', attrs=['bold']))
                                try:
                                    stdin, stdout, stderr = client.exec_command("cd /media/mmcblk0p1/opt/v750/system/ud\n" f"cp File_112_00000 File_112_{SEQUENCIALBKPINT}")
                                    print(f'File_112_00000 - Renomeado para:',end=' ')
                                    print(colored(f'File_112_{SEQUENCIALBKPINT}\n', 'green'))
                                except:
                                    print(colored('NÃO FOI POSSIVEL RENOMEAR File_112_00000.', 'red'))
                                    print(colored('VERIFIQUE O SEQUENCIAL DO ARQUIVO E RENOMEIE MANUALMENTE!\n', 'yellow'))
                                    pass
                                
                            ### CRIANDO PCGAR'S ###
                            print(colored("  --> VERIFICANDO E GERANDO ARQUIVOS PCGAR'S DE BKP <--", 'blue', attrs=['bold']))
                            stdin, stdout, stderr = client.exec_command("cd /media/mmcblk0p1/opt/v750/system/ud\n"f"/root/app/{EXECUTAVEL_A} -R")
                            time.sleep(2)
                            stdin, stdout, stderr = client.exec_command("cd /media/mmcblk0p1/opt/v750/system/ud\n""ls PCGAR_*")
                            PCGAR = stdout.read().decode('utf8')
                            print(PCGAR[:-1])
                                
                            if PCGAR[:5] == 'PCGAR':
                                print(colored('SALVANDO OS ARQUIVOS PCGAR DE BKP EXISTENTES!!!', 'blue', attrs=['bold']))
                                ssh_command(client, "cd /media/mmcblk0p1/opt/v750/system/ud\n""tar -cf PCGAR.tar.gz PCGAR_*")
                                print('DOWNLOAD: ...', end='')
                                time.sleep(1)
                                ### DOWNLOAD DO ARQUIVO ###
                                ftp = client.open_sftp()
                                ftp.get(FILE_PCGAR2, LOCAL_PCGAR)
                                time.sleep(3)
                                print(colored('\rDOWNLOAD: OK ', 'green'))
                                ftp.close()
                                print()
                                INFO_PCGAR_TG = 'PCGAR'
                                SEQUENCIAL_INFO = (str(SEQUENCIALBKPINT))
                            else:
                                print(colored('SEM ARQUIVOS PCGAR PARA RESGATAR!!!\n', 'red', attrs=['bold']))
                                INFO_PCGAR_TG = 'NAO'																
                                time.sleep(2)
                        else:
                            print(colored('PASTA SYSTEM NO SDCARD NAO LOCALIZADA\n-> SDCARD COM DEFEITO!!!\n', 'red'))
                            time.sleep(1)
                            SD_PLAN = 'SDCARD COM DEFEITO'
                            LOCALIZADOR = open(g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\'+g_OS+' - '+g_NS+'.txt', 'a')
                            LOCALIZADOR.write('\n-> SDCARD COM DEFEITO')
                            LOCALIZADOR.close()
                    else:
                        print(colored('--> SDCARD COM DEFEITO!!!\n', 'red'))
                        time.sleep(1)
                        SD_PLAN = 'SDCARD COM DEFEITO'
                        LOCALIZADOR = open(g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\'+g_OS+' - '+g_NS+'.txt', 'a')
                        LOCALIZADOR.write('\n-> SDCARD COM DEFEITO')
                        LOCALIZADOR.close()

                elif EXECUTAVELYY[:-1] == "SPTrans-v3680hf":
                    print(colored('VALIDADOR DE BKP: SPTRANS\n', attrs=['bold']))
                    EXECUTAVEL_A = "SPTrans-v3680hf"
                    time.sleep(2)
                    print(colored('---> INICIANDO BKP DA PASTA SYSTEM EM SDCARD <---', 'blue', attrs=['bold']))
                    stdin, stdout, stderr = client.exec_command('find /media/mmcblk0p1/opt/v750/system/ud')
                    time.sleep(2)
                    try:
                        VERFSD = stdout.read().decode('utf8')
                    except:
                        VERFSD = 'SDCARD COM DEFEITO'
                    
                    if VERFSD[0:35] == "/media/mmcblk0p1/opt/v750/system/ud":
                        stdin, stdout, stderr = client.exec_command('cd /media/mmcblk0p1/opt/v750/system/ud/backup/ \n''ls')
                        VERFSD2 = stdout.read().decode('utf8')
                        if VERFSD2 == "":
                            print(colored('AVISO: Pasta BACKUP do SDCARD está vazia!', 'yellow', attrs=['bold']))
                        else:
                            pass

                        stdin, stdout, stderr = client.exec_command("cd /media/mmcblk0p1/opt/v750/system/ud \n""tar --exclude='photos' -cf System.tar.gz *")
                        time.sleep(10)
                        print(colored('     Arquivo TAR do SDCARD:', attrs=['bold']))
                        ssh_command(client, "cd /media/mmcblk0p1/opt/v750/system/ud \n""ls -lh System*")
                        print('DOWNLOAD: ...', end=' ')
                        time.sleep(2)
                        ### DOWNLOAD DO ARQUIVO ###
                        ftp = client.open_sftp()
                        ftp.get(FILE_SD, LOCAL_SD)
                        ftp.close()
                        print(colored('\rDOWNLOAD: OK  ', 'green', attrs=['bold']))
                        print()

                        print(colored('---> VERIFICANDO ARQUIVOS PCGAR DE BKP <---', 'blue', attrs=['bold']))
                        #ssh_command(client,"cd /opt/v750/system/ud/\n""ls File_002_*\n")
                        stdin, stdout, stderr = client.exec_command("cd /media/mmcblk0p1/opt/v750/system/ud/\n""du -hsb File_112_00000\n")
                        time.sleep(1)
                        KB_00000 = stdout.read().decode('utf8')
                        KB = KB_00000.find('File_112_00000')

                        if KB == -1:
                            print(colored("SEM ARQUIVO DE BKP PARA RESGATAR!!!\n", 'red', attrs=['bold']))
                            time.sleep(2)
                            INFO_PCGAR_TG = 'NAO'
                            
                        elif KB >= 0:
                            if (int(KB_00000[:-15])) <= 1000:
                                print(KB_00000[:-15]+'File_112_00000') 
                                print(colored(' --> File_112_00000 MENOR QUE 1000Kb <--\n', 'red', attrs=['bold']))
                                time.sleep(1.5)
                            else:
                                print(colored(' --> INICIANDO RESGATE DO ARQUIVO File_112_00000 <--', 'green', attrs=['bold']))
                                print(KB_00000[:-15]+'File_112_00000')
                                print(colored('\n -> VERIFICANDO SEQUENCIAL <-',attrs=['bold']))
                                stdin, stdout, stderr = client.exec_command("cd /media/mmcblk0p1/opt/v750/system/ud/backup\n""ls -t File_112_*")
                                time.sleep(2)
                                BACKUP = stdout.read().decode('utf8')
                                SEQUENCIALBKP = (BACKUP[9:14])

                                try:
                                    SEQUENCIALBKPINT = (int(SEQUENCIALBKP)+1)
                                except:
                                    SEQUENCIALBKPINT = (int(g_NS))

                                if (len(str(SEQUENCIALBKPINT))) == 5:
                                    pass
                                elif (len(str(SEQUENCIALBKPINT))) == 4:
                                    SEQUENCIALBKPINT = '0'+(str(SEQUENCIALBKPINT))
                                elif (len(str(SEQUENCIALBKPINT))) == 3:
                                    SEQUENCIALBKPINT = '00'+(str(SEQUENCIALBKPINT))
                                elif (len(str(SEQUENCIALBKPINT))) == 2:
                                    SEQUENCIALBKPINT = '000'+(str(SEQUENCIALBKPINT))
                                elif (len(str(SEQUENCIALBKPINT))) == 1:
                                    SEQUENCIALBKPINT = '0000'+(str(SEQUENCIALBKPINT))

                                print(colored(f'SEQUENCIAL:', attrs=['bold']), end=' ')
                                print(colored(f'{SEQUENCIALBKPINT}', 'green'))
                                time.sleep(1)

                                ### RENOMEANDO FILE_112_00000 ###
                                print(colored("\n    --> RENOMEANDO FILE 112 <--", 'blue', attrs=['bold']))
                                try:
                                    stdin, stdout, stderr = client.exec_command("cd /media/mmcblk0p1/opt/v750/system/ud\n" f"cp File_112_00000 File_112_{SEQUENCIALBKPINT}")
                                    print(f'File_112_00000 - Renomeado para:',end=' ')
                                    print(colored(f'File_112_{SEQUENCIALBKPINT}\n', 'green'))
                                except:
                                    print(colored('NÃO FOI POSSIVEL RENOMEAR File_112_00000.', 'red'))
                                    print(colored('VERIFIQUE O SEQUENCIAL DO ARQUIVO E RENOMEIE MANUALMENTE!\n', 'yellow'))
                                    pass
                                
                            ### CRIANDO PCGAR'S ###
                            print(colored("  --> VERIFICANDO E GERANDO ARQUIVOS PCGAR'S DE BKP <--", 'blue', attrs=['bold']))
                            stdin, stdout, stderr = client.exec_command("cd /media/mmcblk0p1/opt/v750/system/ud\n"f"/root/app/{EXECUTAVEL_A} -R")
                            time.sleep(2)
                            stdin, stdout, stderr = client.exec_command("cd /media/mmcblk0p1/opt/v750/system/ud\n""ls PCGAR_*")
                            PCGAR = stdout.read().decode('utf8')
                            print(PCGAR[:-1])
                                
                            if PCGAR[:5] == 'PCGAR':
                                print(colored('SALVANDO OS ARQUIVOS PCGAR DE BKP EXISTENTES!!!', 'blue', attrs=['bold']))
                                ssh_command(client, "cd /media/mmcblk0p1/opt/v750/system/ud\n""tar -cf PCGAR.tar.gz PCGAR_*")
                                print('DOWNLOAD: ...', end='')
                                time.sleep(1)
                                ### DOWNLOAD DO ARQUIVO ###
                                ftp = client.open_sftp()
                                ftp.get(FILE_PCGAR2, LOCAL_PCGAR)
                                time.sleep(3)
                                print(colored('\rDOWNLOAD: OK ', 'green'))
                                ftp.close()
                                print()
                                INFO_PCGAR_TG = 'PCGAR'
                                SEQUENCIAL_INFO = (str(SEQUENCIALBKPINT))
                            else:
                                print(colored('SEM ARQUIVOS PCGAR PARA RESGATAR!!!\n', 'red', attrs=['bold']))
                                INFO_PCGAR_TG = 'NAO'																
                                time.sleep(2)
                        else:
                            print(colored('PASTA SYSTEM NO SDCARD NAO LOCALIZADA\n-> SDCARD COM DEFEITO!!!\n', 'red'))
                            time.sleep(1)
                            SD_PLAN = 'SDCARD COM DEFEITO'
                            LOCALIZADOR = open(g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\'+g_OS+' - '+g_NS+'.txt', 'a')
                            LOCALIZADOR.write('\n-> SDCARD COM DEFEITO')
                            LOCALIZADOR.close()
                    else:
                        print(colored('--> SDCARD COM DEFEITO!!!\n', 'red'))
                        time.sleep(1)
                        SD_PLAN = 'SDCARD COM DEFEITO'
                        LOCALIZADOR = open(g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\'+g_OS+' - '+g_NS+'.txt', 'a')
                        LOCALIZADOR.write('\n-> SDCARD COM DEFEITO')
                        LOCALIZADOR.close()
                
                
                elif EXECUTAVELZ[:-1] == "Mercury-v3695":
                    print(colored('VALIDADOR DE BKP: TOP / PROJETOS\n', attrs=['bold']))
                    EXECUTAVEL_A = "Mercury-v3695"
                    time.sleep(2)
                    print(colored('---> INICIANDO BKP DA PASTA SYSTEM EM SDCARD <---', 'blue', attrs=['bold']))
                    stdin, stdout, stderr = client.exec_command('find /media/mmcblk0p1/opt/v750/system/ud')
                    time.sleep(2)
                    try:
                        VERFSD = stdout.read().decode('utf8')
                    except:
                        VERFSD = 'SDCARD COM DEFEITO'
                    
                    if VERFSD[0:35] == "/media/mmcblk0p1/opt/v750/system/ud":
                        stdin, stdout, stderr = client.exec_command("cd /media/mmcblk0p1/\n""tar -cf mmcblk0p1.tar.gz *")
                        time.sleep(2)
                        print(colored('     Arquivo TAR do SDCARD:', attrs=['bold']))
                        ssh_command(client, "cd /media/mmcblk0p1/\n""ls -lh mmcblk0p1*")
                        print('DOWNLOAD: ...', end='')
                        time.sleep(1)
                        ### DOWNLOAD DO ARQUIVO ###
                        try:    
                            ftp = client.open_sftp()
                            ftp.get(FILE_SD_TOP, LOCAL_SD_TOP)
                            time.sleep(3)
                            print(colored('\rDOWNLOAD: OK ', 'green'))
                            ftp.close()
                        except:
                            print(colored('\rDOWNLOAD: FALHA (POSSÍVEL ARQUIVO CORRIMPIDO) ', 'red'))
                            ftp.close()
                            pass
                        print()

                        print(colored('---> VERIFICANDO ARQUIVOS TG-TOP DE BKP <---', 'blue', attrs=['bold']))
                        stdin, stdout, stderr = client.exec_command("cd /media/mmcblk0p1/opt/v750/system/ud/\n""du -hsb BKP_File_002_00000\n")
                        KB_00000 = stdout.read().decode('utf8')
                        KB = KB_00000.find('BKP_File_002_00000')

                        if KB == -1:
                            print(colored("SEM ARQUIVOS TG-TOP PARA RESGATAR!!!\n", 'red'))
                            time.sleep(1.5)
                            INFO_PCGAR_TG = 'NAO'
                        
                        elif KB >= 0:
                            if (int(KB_00000[:-19])) <= 8000:
                                print(KB_00000[:-19]+'BKP_File_002_00000')
                                print(colored(' --> BKP_File_002_00000 MENOR QUE 8000Kb <--\n', 'red'))
                                time.sleep(1.5)
                                INFO_PCGAR_TG = 'NAO'
                            else:
                                print(colored('\n    --> INICIANDO RESGATE DO ARQUIVO BKP_File_002_00000 <--', 'green', attrs=['bold']))
                                stdin, stdout, stderr = client.exec_command("cd /media/mmcblk0p1/opt/v750/system/ud/\n"f"/root/app/{EXECUTAVEL_A} -R")
                                time.sleep(2)
                                stdin, stdout, stderr = client.exec_command("cd /media/mmcblk0p1/opt/v750/system/ud/\n""ls TG*")
                                TG = stdout.read().decode('utf8')
                                        
                                print(colored("\nSALVANDO OS BKPS TG-TOP EXISTENTES!!!", 'blue', attrs=['bold']))
                                ssh_command(client, "cd /media/mmcblk0p1/opt/v750/system/ud\n""tar -cf TG.tar.gz TG*")
                                print('DOWNLOAD: ...', end='')
                                time.sleep(2)
                                ### DOWNLOAD DO ARQUIVO ###
                                ftp = client.open_sftp()
                                ftp.get(FILE_TG2, LOCAL_TG)
                                time.sleep(3)
                                print(colored('\rDOWNLOAD: OK ', 'green'))
                                ftp.close()
                                INFO_PCGAR_TG = 'TOP'
                                SEQUENCIAL_INFO = TG[18:23]

                        ### VERIFICANDO ARQUIVOS TG-BOM ###
                        print(colored("--> VERIFICANDO ARQUIVOS TG'S BOM DE BKP <--", 'blue', attrs=['bold']))
                        stdin, stdout, stderr = client.exec_command("cd /media/mmcblk0p1/001/opt/v750/system/ud/\n""du -hsb File_002_*\n")
                        KB_00001 = stdout.read().decode('utf8')
                        KB2 = KB_00001.find('BKP_File_002_00000') 
                        
                        if KB2 == -1:
                            print(colored("SEM ARQUIVOS TG PARA RESGATAR!!!\n", 'red'))
                            time.sleep(1.5)
                        
                        elif KB2 >= 0:
                            if (int(KB_00001[:-19])) <= 8000:
                                print(KB_00001[:-19]+'File_002_00000')
                                print(colored(' --> BKP_File_002_00000 MENOR QUE 8000Kb <--\n', 'red'))
                                time.sleep(1.5)
                            else:
                                stdin, stdout, stderr = client.exec_command("cd /media/mmcblk0p1/001/opt/v750/system/ud\n"f"/root/app/{EXECUTAVEL_A} -R")
                                time.sleep(2)
                                stdin, stdout, stderr = client.exec_command("cd /media/mmcblk0p1/001/opt/v750/system/ud\n""ls TG*")
                                TG_BOM = stdout.read().decode('utf8')
                                            
                                print(colored("\nSALVANDO OS ARQUIVOS TG-BOM EXISTENTES!!!", 'blue', attrs=['bold']))
                                ssh_command(client, "cd /media/mmcblk0p1/001/opt/v750/system/ud\n""tar -cf TG-BOM.tar.gz TG*")
                                print('DOWNLOAD: ...', end='')
                                time.sleep(2)
                                ### DOWNLOAD DO ARQUIVO ###
                                ftp = client.open_sftp()
                                ftp.get(FILE_TG_BOM2, LOCAL_TG_BOM)
                                time.sleep(5)
                                print(colored('\rDOWNLOAD: OK ', 'green'))
                                ftp.close()
                                INFO_PCGAR_BOM = ' / BOM'
                                SEQUENCIAL_INFO2 = ' /'+TG_BOM[18:23]
                    else:
                        print(colored('--> SDCARD COM DEFEITO!!!\n', 'red'))
                        time.sleep(1)
                        SD_PLAN = 'SDCARD COM DEFEITO'
                        LOCALIZADOR = open(g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\'+g_OS+' - '+g_NS+'.txt', 'a')
                        LOCALIZADOR.write('\n-> SDCARD COM DEFEITO')
                        LOCALIZADOR.close()

                else:
                    print(colored('---> VALIDADOR FALTANDO APLICAÇÃO OU APP NAO IDENTIFICADO!!! <---', 'red'))

            RESGATE = load_workbook(g_C_D+f'\\BKP\\RESGATE-{ANO}.xlsx')
            PLANILHA = RESGATE['INTERVENCAO']
            try:
                ### ADCIONINANDO INFORMACOES NA PLANILHA ###
                RESGATE_INFO = ['F', g_EMPRESA, g_OS, g_NS, INFO_PCGAR_TG+INFO_PCGAR_BOM, SEQUENCIAL_INFO + SEQUENCIAL_PCGAR + SEQUENCIAL_INFO2, DATAXLSX, BKP_PLAN, SD_PLAN + OBS_PLAN]#.alignment = Alignment(horizontal='center')
                PLANILHA.append(RESGATE_INFO)
                NLINHA = PLANILHA.max_row
                ULTIMA_LINHA = [f'A{NLINHA}', f'B{NLINHA}', f'C{NLINHA}', f'D{NLINHA}', f'E{NLINHA}', f'F{NLINHA}', f'G{NLINHA}', f'H{NLINHA}', f'I{NLINHA}']
                ULTIMA_LINHA2 = [f'B{NLINHA}', f'C{NLINHA}', f'D{NLINHA}', f'E{NLINHA}', f'F{NLINHA}', f'G{NLINHA}', f'H{NLINHA}', f'I{NLINHA}']
                #BORDA = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                for cell in ULTIMA_LINHA2:
                    PLANILHA[cell].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                for cell in ULTIMA_LINHA:
                    PLANILHA[cell].alignment = Alignment(horizontal='center')
                RESGATE.save(g_C_D+f'\\BKP\\RESGATE-{ANO}.xlsx')
                print(colored('\n   GRAVADA INFORMAÇÕES NA PLANILHA!!!', 'green'))    
            except:
                print(colored('\nFALHA AO GRAVAR INFORMAÇÕES NA PLANILHA!!!', 'red'))
                print(colored('VERIFIQUE SE A PLANILHA ESTA ABERTA E FECHEA.', 'yellow'))
                pass

            ###### LIMPANDO TARS DO VALIDADOR #####
            print(colored("\n--> LIMPANDO RESGISTROS DO VALIDADOR <--", 'yellow', attrs=['bold']))
            print("LIMPANDO...", end='')
            try:
                stdin, stdout, stderr = client.exec_command(f"rm {FILE_SD}")
                time.sleep(0.5)
                stdin, stdout, stderr = client.exec_command(f"rm {FILE_PCGAR2}")
                time.sleep(0.5)
                stdin, stdout, stderr = client.exec_command(f"rm {FILE_TG_BOM2}")
                time.sleep(0.5)
                stdin, stdout, stderr = client.exec_command(f"rm {FILE_SD_TOP}")
                time.sleep(0.5)
                stdin, stdout, stderr = client.exec_command(f"rm {FILE_TG2}")
                time.sleep(0.5)
                print(colored("\rLIMPEZA CONCLUÍDA !!!", 'green'))
            except:
                print(colored("\rLIMPEZA CONCLUÍDA !!!", 'green'))

            print(colored("\n     --> FECHANDO CONEXAO SSH. <--\n\n", 'red', attrs=['bold']))
            client.close()

            print(colored('---> INCIANDO EXTRAÇÃO DOS ARQUIVOS DE BKP <---', 'blue',attrs=['bold']))
            
            ########################## VERIFICANDO SE EXISTE SDCARD PARA EXTRAÇÃO ##########################
            DESC_SD = (g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\System\\')
            DESC_SD_TOP = (g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\mmcblk0p1\\')
            DESC_TG_BOM = (g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\TG-BOM\\')
            
            ######################## SDCARD ########################
            if(os.path.exists(LOCAL_SD)):
                if os.path.isdir(g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\System'):
                    print(colored("Ja existe uma pasta para a System/SDCARD", 'yellow'))
                else:
                    os.mkdir(g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\System')
                    print(colored("Pasta para a System/SDCARD criada com sucesso!", 'yellow'))
                    
                print(colored('EXTRAINDO ARQUIVOS BKP SDCARD...', attrs=['bold']))
                SDD_LOCAL = tarfile.open(LOCAL_SD)
                try:
                    SDD_LOCAL.extractall(DESC_SD)
                except:
                    pass
                SDD_LOCAL.close()
                print(colored('EXTRAÇÃO COMPLETA - OK\n', 'green'))
                time.sleep(1)
                os.remove(LOCAL_SD)

            elif(os.path.exists(LOCAL_SD_TOP)):
                if os.path.isdir(g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\mmcblk0p1'):
                    print(colored("Ja existe uma pasta para a System/SDCARD", 'yellow'))
                else:
                    os.mkdir(g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\mmcblk0p1')
                    print(colored("Pasta para a System/SDCARD criada com sucesso!", 'yellow'))
                    
                print(colored('EXTRAINDO ARQUIVOS BKP SDCARD...', attrs=['bold']))
                SD_TOP = tarfile.open(LOCAL_SD_TOP)
                try:
                    SD_TOP.extractall(DESC_SD_TOP)
                except:
                    pass
                SD_TOP.close()
                print(colored('EXTRAÇÃO COMPLETA - OK\n', 'green'))
                time.sleep(1)
                os.remove(LOCAL_SD_TOP)

            else:
                print(colored("NAO HA ARQUIVOS DE BKP DO SDCARD PARA EXTRAÇÃO...\n", 'red'))
                time.sleep(1)
            
            ######################## PCGAR ########################
            if (os.path.exists(LOCAL_PCGAR)):
                print(colored('EXTRAINDO ARQUIVOS PCGAR...', attrs=['bold']))
                PCGAR_LOCAL = tarfile.open(LOCAL_PCGAR)
                PCGAR_LOCAL.extractall(PASTA_INI)
                PCGAR_LOCAL.close()
                print(colored('EXTRAÇÃO COMPLETA - OK\n', 'green'))
                time.sleep(1)
                os.remove(LOCAL_PCGAR)
            else:
                print(colored("NAO HA ARQUIVOS PCGAR DE BKP PARA EXTRAÇÃO...\n", 'red'))
                time.sleep(1)

            ######################## TG-TOP ########################
            if (os.path.exists(LOCAL_TG)):
                print(colored('EXTRAINDO ARQUIVOS DE BKP TG-TOP...', attrs=['bold']))
                TG_LOCAL = tarfile.open(LOCAL_TG)
                try:
                    TG_LOCAL.extractall(PASTA_INI)
                    TG_LOCAL.close()
                    os.remove(LOCAL_TG)
                except:
                    pass
                print(colored('EXTRAÇÃO COMPLETA - OK\n', 'green'))
                time.sleep(1)
                print()
            else:
                print(colored("NAO HA ARQUIVOS BKP TG-TOP PARA EXTRAÇÃO...\n", 'red'))
                time.sleep(1)

            ######################## TG-BOM ########################
            if (os.path.exists(LOCAL_TG_BOM)):
                if os.path.isdir(g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\TG-BOM'):
                    print ('Ja existe uma pasta para TG-BOM')
                else:
                    os.mkdir(g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\TG-BOM')
                    print ('Pasta para TG-BOM criada com sucesso!')
                
                print(colored('EXTRAINDO ARQUIVOS DE BKP TG-BOM...', attrs=['bold']))
                TG_BOM_LOCAL = tarfile.open(LOCAL_TG_BOM)
                try:
                    TG_BOM_LOCAL.extractall(DESC_TG_BOM)
                except:
                    pass
                TG_BOM_LOCAL.close()
                print(colored('EXTRAÇÃO COMPLETA - OK\n', 'green'))
                time.sleep(1)
                print()
                os.remove(LOCAL_TG_BOM)
            else:
                print(colored("NAO HA ARQUIVOS TG-BOM DE BKP PARA EXTRAÇÃO...\n", 'red'))
                time.sleep(1)
            
            ######################## COPIAR PARA O SERVIDOR ########################
            print(colored('\n\n  --> COPIANDO PARA O SERVIDOR <--', 'blue', attrs=['bold']))
            os.system("net use J: \\\\172.18.10.2\\assistencia$")
            if os.path.exists(f"J:\#ARQUIVOS#\{g_SP_CMT2}"):
                LOCAL = g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA
                SERV = 'J:\\#ARQUIVOS#\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA
                os.listdir(LOCAL)
                print('COPIANDO...', end=' ')
                try:
                    shutil.copytree(LOCAL, SERV)
                except:
                    pass
                print(colored('\rBKP COM SUCESSO!!! - SERVIDOR (172.18.10.2)', 'green'))
                time.sleep(1)
                os.system('net use J: /delete')
            else:
                print(colored("SEM CONEXÃO COM O SERVIDOR (172.18.10.2)", 'red'))
                os.system('net use J: /delete')

        ######################## EXIBIR ARQUIVOS EXTRAIDOS ########################
            try:
                TG_PCGAR = os.listdir(PASTA_INI)
                for f in TG_PCGAR:
                    f_path = os.path.join(PASTA_INI,f)
                    if(os.path.isfile(f_path)):
                        f_size = os.path.getsize(f_path)
                        print(f_size, f_path)
            except:
                pass
            
            ###################### ABRIR A PASTA ##########################
            os.system(f"explorer.exe "+PASTA_INI)

            ### VERIFICAR SE REINICIA OU NAO ###
            print(colored("\n\nREINICIAR?", 'yellow'))
            REINICIARBKP = input('Digite S para SIM ou ENTER para Fechar...\n').upper().strip()
            if REINICIARBKP == 'S':
                print('REINICIANDO...')
                os.system('python RESGATE.py')
            elif REINICIARBKP == 'SIM':
                print('REINICIANDO...')
                os.system('python RESGATE.py')
            elif REINICIARBKP == 'SS':
                print('REINICIANDO...')
                os.system('python RESGATE.py')
            else:
                print('FECHANDO...\n')
                os.system('TASKKILL /IM python.exe /T /F')
            
    else:
        BKP_PLAN = ' '
        pass

    ############################## VERIFICAR SE É DE FABRICA ##############################
    # EXECUTAVEIS = ['SPTrans-v3068', 'SPTrans-v3680hf', 'SPTrans-v3690w-hf', 'Mercury-v3680hf', 'Control-v3695', 'Mercury-v3695', 'Mercury-v3690hf', 'Mercury-v3680']
    print(colored(f'---> VARIFICANDO APLICAÇÃO DO VALIDADOR <---', 'blue', attrs=['bold']))
    stdin, stdout, stderr = client.exec_command('cd /root/app\n''ls SPTrans-v3068')
    time.sleep(0.5)
    EXECUTAVEL1 = stdout.read().decode('utf8')

    stdin, stdout, stderr = client.exec_command('cd /root/app\n''ls SPTrans-v3680hf')
    time.sleep(0.5)
    EXECUTAVEL2 = stdout.read().decode('utf8')

    stdin, stdout, stderr = client.exec_command('cd /root/app\n''ls SPTrans-v3690w-hf')
    time.sleep(0.5)
    EXECUTAVEL3 = stdout.read().decode('utf8')

    stdin, stdout, stderr = client.exec_command('cd /root/app\n''ls Mercury-v3680hf')
    time.sleep(0.5)
    EXECUTAVEL4 = stdout.read().decode('utf8')

    stdin, stdout, stderr = client.exec_command('cd /root/app\n''ls Control-v3695')
    time.sleep(0.5)
    EXECUTAVEL5 = stdout.read().decode('utf8')

    stdin, stdout, stderr = client.exec_command('cd /root/app\n''ls Mercury-v3695')
    time.sleep(0.5)
    EXECUTAVEL6 = stdout.read().decode('utf8')

    stdin, stdout, stderr = client.exec_command('cd /root/app\n''ls Mercury-v3690hf')
    time.sleep(0.5)
    EXECUTAVEL7 = stdout.read().decode('utf8')

    stdin, stdout, stderr = client.exec_command('cd /root/app\n''ls Mercury-v3680')
    time.sleep(0.5)
    EXECUTAVEL8 = stdout.read().decode('utf8')

    stdin, stdout, stderr = client.exec_command('cd /root/app\n''ls Mercury-v3690')
    time.sleep(0.5)
    EXECUTAVEL9 = stdout.read().decode('utf8')

    if EXECUTAVEL1[:-1] == "SPTrans-v3068":
            print(colored('APP: SPTrans-v3068\n', attrs=['bold']))
            EXECUTAVELX = "SPTrans-v3068"
            time.sleep(1)

    elif EXECUTAVEL2[:-1] == "SPTrans-v3680hf":
            print(colored('APP: SPTrans-v3680hf\n', attrs=['bold']))
            EXECUTAVELX = "SPTrans-v3680hf"
            time.sleep(1)
        
    elif EXECUTAVEL3[:-1] == "SPTrans-v3690w-hf":
            print(colored('APP: SPTrans-v3690w-hf\n', attrs=['bold']))
            EXECUTAVELX = "SPTrans-v3690w-hf"
            time.sleep(1)

    elif EXECUTAVEL4[:-1] == "Mercury-v3680hf":
            print(colored('APP: Mercury-v3680hf\n', attrs=['bold']))
            EXECUTAVELX = "Mercury-v3680hf"
            time.sleep(1)
        
    elif EXECUTAVEL5[:-1] == "Control-v3695":
            print(colored('APP: Control-v3695\n', attrs=['bold']))
            EXECUTAVELX = "Control-v3695"
            time.sleep(1)
        
    elif EXECUTAVEL6[:-1] == "Mercury-v3695":
            print(colored('APP: Mercury-v3695\n', attrs=['bold']))
            EXECUTAVELX = "Mercury-v3695"
            time.sleep(1)

    elif EXECUTAVEL7[:-1] == "Mercury-v3690hf":
            print(colored('APP: Mercury-v3690hf\n', attrs=['bold']))
            EXECUTAVELX = "Mercury-v3690hf"
            time.sleep(1)
        
    elif EXECUTAVEL8[:-1] == "Mercury-v3680":
            print(colored('APP: Mercury-v3680\n', attrs=['bold']))
            EXECUTAVELX = "Mercury-v3680"
            time.sleep(1)
        
    elif EXECUTAVEL9[:-1] == "Mercury-v3690":
            print(colored('APP: Mercury-v3690\n', attrs=['bold']))
            EXECUTAVELX = "Mercury-v3690"
            time.sleep(1)

    else:
            print(colored('!!! EQUIPAMENTO COM SOFTWARE DE FABRICA !!!', 'red', attrs=['bold']))
            print(colored('Nada a resgatar...', attrs=['bold']))
            print(colored("\n\n     --> FECHANDO CONEXAO SSH. <--", 'yellow'))
            client.close()

            LOCALIZADOR = open(g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\'+g_OS+' - '+g_NS+'.txt', 'a')
            LOCALIZADOR.write('\n-> EQUIPAMENTO COM SOFTWARE DE FABRICA')
            LOCALIZADOR.close()
            
            RESGATE = load_workbook(g_C_D+f'\\BKP\\RESGATE-{ANO}.xlsx')
            PLANILHA = RESGATE['INTERVENCAO']
            try:
                ### ADCIONINANDO INFORMACOES NA PLANILHA ###
                RESGATE_INFO = ['F', g_EMPRESA, g_OS, g_NS, 'NAO', ' ', DATAXLSX, 'EQUIPAMENTO DE FABRICA']
                PLANILHA.append(RESGATE_INFO)
                NLINHA = PLANILHA.max_row
                ULTIMA_LINHA = [f'A{NLINHA}', f'B{NLINHA}', f'C{NLINHA}', f'D{NLINHA}', f'E{NLINHA}', f'F{NLINHA}', f'G{NLINHA}', f'H{NLINHA}', f'I{NLINHA}']
                ULTIMA_LINHA2 = [f'B{NLINHA}', f'C{NLINHA}', f'D{NLINHA}', f'E{NLINHA}', f'F{NLINHA}', f'G{NLINHA}', f'H{NLINHA}', f'I{NLINHA}']
                #BORDA = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                for cell in ULTIMA_LINHA2:
                    PLANILHA[cell].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                for cell in ULTIMA_LINHA:
                    PLANILHA[cell].alignment = Alignment(horizontal='center')
                RESGATE.save(g_C_D+f'\\BKP\\RESGATE-{ANO}.xlsx')
                print(colored('\n   GRAVADA INFORMAÇÕES NA PLANILHA!!!', 'green'))
            except:
                print(colored('\nFALHA AO GRAVAR INFORMAÇÕES NA PLANILHA!!!', 'red'))
                print(colored('VERIFIQUE SE A PLANILHA ESTA ABERTA E FECHEA.', 'yellow'))
                pass

            ### VERIFICAR SE REINICIA OU NAO ###
            print(colored("\n\nREINICIAR?", 'yellow'))
            REINICIARFABRICA = input('Digite S para SIM ou ENTER para Fechar...\n').upper()
            if REINICIARFABRICA == 'S':
                print('REINICIANDO...')
                sys.stdout.flush()
                os.system('python RESGATE.py')
            elif REINICIARFABRICA == 'SIM':
                print('REINICIANDO...')
                sys.stdout.flush()
                os.system('python RESGATE.py')
            elif REINICIARFABRICA == 'SS':
                print('REINICIANDO...')
                sys.stdout.flush()
                os.system('python RESGATE.py')
            else:
                print('FECHANDO...\n')
                os.system('TASKKILL /IM python.exe /T /F')

    ############# DEFININDO QUAL TIPO DE EMPRESA (1 - SPTRANS / 2 - CMT-PROJETO / 3 - TOP-BOM) ##############
        
    ############################################## SPTRANS ##################################################
    if g_SP_CMT2 == 'SPTRANS':
            
        ######################## INFO PLANILHA PARA NAO FALHAR ########################
        OBS_PLAN = ' '
        INFO_PCGAR_BOM = ' '
        SEQUENCIAL_BOM = ' '
        SERIAL_PLAN = ' '
        
        ### VERIFICANDO NUMERO DE SERIE INTERNO ###
        print(colored("---> VERIFICANDO NUMERO DE SERIE <---", 'blue', attrs=['bold']))
        stdin, stdout, stderr = client.exec_command('ls serial.txt')
        SERIAL = stdout.read().decode('utf8')
        SERIAL_SERIALREADER = SERIAL.find("serial.txt")

        if SERIAL_SERIALREADER >= 0:
                stdin, stdout, stderr = client.exec_command("tail -1 serial.txt\n")
                SERIAL = stdout.read().decode('utf8')

                if SERIAL == "":
                    print(colored('Serial do Validador: ARQUIVO TXT SEM INFORMAÇÃO', attrs=['bold']))
                    print('')
                    ### DOWNLOAD DO TXT ###
                    ftp = client.open_sftp()
                    ftp.get(FILE_TXT, PASTA_TXT)
                    ftp.close()
                elif SERIAL[:5] == g_NS:
                    print(colored('Serial do Validador:', attrs=['bold']), end=' ')
                    print(SERIAL[:5])
                    print(colored('SERIAL COMPATIVEL - OK', 'green'))
                    print('')
                    ### DOWNLOAD DO TXT ###
                    ftp = client.open_sftp()
                    ftp.get(FILE_TXT, PASTA_TXT)
                    ftp.close()
                elif '0'+SERIAL[:4] == g_NS:
                    print(colored('Serial do Validador:', attrs=['bold']), end=' ')
                    print(SERIAL[:4])
                    print(colored('SERIAL COMPATIVEL - OK', 'green'))
                    print('')
                    ### DOWNLOAD DO TXT ###
                    ftp = client.open_sftp()
                    ftp.get(FILE_TXT, PASTA_TXT)
                    ftp.close()
                elif SERIAL[:4] == g_NS:
                    print(colored('Serial do Validador:', attrs=['bold']), end=' ')
                    print(SERIAL[:4])
                    print(colored('SERIAL COMPATIVEL - OK', 'green'))
                    print('')
                    ### DOWNLOAD DO TXT ###
                    ftp = client.open_sftp()
                    ftp.get(FILE_TXT, PASTA_TXT)
                    ftp.close()
                elif SERIAL[:3] == g_NS:
                    print(colored('Serial do Validador:', attrs=['bold']), end=' ')
                    print(SERIAL[:3])
                    print(colored('SERIAL COMPATIVEL - OK', 'green'))
                    print('')
                    ### DOWNLOAD DO TXT ###
                    ftp = client.open_sftp()
                    ftp.get(FILE_TXT, PASTA_TXT)
                    ftp.close()
                elif SERIAL[:2] == g_NS:
                    print(colored('Serial do Validador:', attrs=['bold']), end=' ')
                    print(SERIAL[:2])
                    print(colored('SERIAL COMPATIVEL - OK', 'green'))
                    print('')
                    ### DOWNLOAD DO TXT ###
                    ftp = client.open_sftp()
                    ftp.get(FILE_TXT, PASTA_TXT)
                    ftp.close()
                elif SERIAL[:1] == g_NS:
                    print(colored('Serial do Validador:', attrs=['bold']), end=' ')
                    print(SERIAL[:1])
                    print(colored('SERIAL COMPATIVEL - OK', 'green'))
                    print('')
                    ### DOWNLOAD DO TXT ###
                    ftp = client.open_sftp()
                    ftp.get(FILE_TXT, PASTA_TXT)
                    ftp.close()
                else:
                    if '0' in SERIAL:
                        print(colored('DIVERGENTE DO BACKSHELL:', 'red', attrs=['bold']))
                        print(colored('Serial OS:', attrs=['bold']), end=' ')
                        print(g_NS)
                        print(colored('Serial Validador:', attrs=['bold']), end=' ')
                        print(SERIAL[:5])
                        print('')
                        ### DOWNLOAD DO TXT ###
                        ftp = client.open_sftp()
                        ftp.get(FILE_TXT, PASTA_TXT)
                        ftp.close()
                        SERIAL_PLAN = ('NS: '+SERIAL[:5])

        else:
                stdin, stdout, stderr = client.exec_command('ls serialreader.txt')
                SERIAL2 = stdout.read().decode('utf8')
                SERIAL_READER = SERIAL2.find("serialreader.txt")
                
                if SERIAL_READER >= 0:
                    stdin, stdout, stderr = client.exec_command("tail -1 serialreader.txt\n")
                    SERIALREAD = stdout.read().decode('utf8')

                    if SERIALREAD == "":
                        print(colored('Serial do Validador: ARQUIVO TXT SEM INFORMAÇÃO', attrs=['bold']))
                        print('')
                        ### DOWNLOAD DO TXT ###
                        ftp = client.open_sftp()
                        ftp.get(FILE_TXT_CMT, PASTA_TXT_CMT)
                        ftp.close()
                    elif SERIALREAD[:5] == g_NS:
                        print(colored('Serial do Validador:', attrs=['bold']), end=' ')
                        print(SERIALREAD[:5])
                        print(colored('SERIAL COMPATIVEL - OK', 'green'))
                        print('')
                        ### DOWNLOAD DO TXT ###
                        ftp = client.open_sftp()
                        ftp.get(FILE_TXT_CMT, PASTA_TXT_CMT)
                        ftp.close()
                    elif '0'+SERIALREAD[:4] == g_NS:
                        print(colored('Serial do Validador:', attrs=['bold']), end=' ')
                        print(SERIALREAD[:4])
                        print(colored('SERIAL COMPATIVEL - OK', 'green'))
                        print('')
                        ### DOWNLOAD DO TXT ###
                        ftp = client.open_sftp()
                        ftp.get(FILE_TXT_CMT, PASTA_TXT_CMT)
                        ftp.close()
                    elif SERIALREAD[:4] == g_NS:
                        print(colored('Serial do Validador:', attrs=['bold']), end=' ')
                        print(SERIALREAD[:4])
                        print(colored('SERIAL COMPATIVEL - OK', 'green'))
                        print('')
                        ### DOWNLOAD DO TXT ###
                        ftp = client.open_sftp()
                        ftp.get(FILE_TXT_CMT, PASTA_TXT_CMT)
                        ftp.close()
                    elif SERIALREAD[:3] == g_NS:
                        print(colored('Serial do Validador:', attrs=['bold']), end=' ')
                        print(SERIALREAD[:3])
                        print(colored('SERIAL COMPATIVEL - OK', 'green'))
                        print('')
                        ### DOWNLOAD DO TXT ###
                        ftp = client.open_sftp()
                        ftp.get(FILE_TXT_CMT, PASTA_TXT_CMT)
                        ftp.close()
                    elif SERIALREAD[:2] == g_NS:
                        print(colored('Serial do Validador:', attrs=['bold']), end=' ')
                        print(SERIALREAD[:2])
                        print(colored('SERIAL COMPATIVEL - OK', 'green'))
                        print('')
                        ### DOWNLOAD DO TXT ###
                        ftp = client.open_sftp()
                        ftp.get(FILE_TXT_CMT, PASTA_TXT_CMT)
                        ftp.close()
                    elif SERIALREAD[:1] == g_NS:
                        print(colored('Serial do Validador:', attrs=['bold']), end=' ')
                        print(SERIALREAD[:1])
                        print(colored('SERIAL COMPATIVEL - OK', 'green'))
                        print('')
                        ### DOWNLOAD DO TXT ###
                        ftp = client.open_sftp()
                        ftp.get(FILE_TXT_CMT, PASTA_TXT_CMT)
                        ftp.close()
                    else:
                        if '0' in SERIALREAD:
                            print(colored('DIVERGENTE DO BACKSHELL:', 'red', attrs=['bold']))
                            print(colored('Serial OS:', attrs=['bold']), end=' ')
                            print(g_NS)
                            print(colored('Serial Validador:', attrs=['bold']), end=' ')
                            print(SERIALREAD[:5])
                            print('')
                            ### DOWNLOAD DO TXT ###
                            ftp = client.open_sftp()
                            ftp.get(FILE_TXT_CMT, PASTA_TXT_CMT)
                            ftp.close()
                            SERIAL_PLAN = ('NS: '+SERIALREAD[:5])
                else:
                    print(colored('ARQUIVO TXT NAO ENCONTRADO!', 'red', attrs=['bold']))
                    print('')

        ### VERIFICANDO A EXISTENCIA DO SDCARD ###
        print(colored("---> VERIFICANDO SDCARD <---", 'blue', attrs=['bold']))
        time.sleep(1)
        stdin, stdout, stderr = client.exec_command("fdisk -l /dev/mmcblk0p1")
        SDCARD = stdout.read().decode('utf8')
        SD = SDCARD.find('/dev/mmcblk0p1')
        print(SDCARD[1:29], end='')
            
        if SD == -1:
                ### SDCARD INOPERANTE OU NAO EXISTE ###
                print(colored('\r   !!! SDCard não detectado !!!\n', 'red'))
                time.sleep(1)
                SD_PLAN = 'SDCARD COM DEFEITO'

                LOCALIZADOR = open(g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\'+g_OS+' - '+g_NS+'.txt', 'a')
                LOCALIZADOR.write('\n-> SDCARD COM DEFEITO')
                LOCALIZADOR.close()
        else:
                ### INICIANDO BKP DA SYSTEM EM SDCARD ###
                print(colored('\n    !!! SDCard - OK !!!\n', 'green'))
                SD_PLAN = ' '
                time.sleep(1)
                print(colored("---> INICIANDO BKP DA PASTA SYSTEM EM SDCARD <---", 'blue', attrs=['bold']))
                stdin, stdout, stderr = client.exec_command('find /media/mmcblk0p1/opt/v750/system/ud')
                time.sleep(2)
                try:
                    VERFSD = stdout.read().decode('utf8')
                except:
                    VERFSD = 'SDCARD COM DEFEITO'
                
                if VERFSD[0:35] == "/media/mmcblk0p1/opt/v750/system/ud":
                    stdin, stdout, stderr = client.exec_command('cd /media/mmcblk0p1/opt/v750/system/ud/backup/ \n''ls')
                    VERFSD2 = stdout.read().decode('utf8')
                    if VERFSD2 == "":
                        print(colored('AVISO: Pasta BACKUP do SDCARD está vazia!', 'yellow', attrs=['bold']))
                    else:
                        pass
                    stdin, stdout, stderr = client.exec_command("cd /media/mmcblk0p1/opt/v750/system/ud \n""tar --exclude='photos' -cf System.tar.gz *")
                    time.sleep(8)
                    print(colored('Arquivo TAR do SDCARD:', attrs=['bold']))
                    ssh_command(client, "cd /media/mmcblk0p1/opt/v750/system/ud \n""ls -lh System*")
                    print('DOWNLOAD: ...', end='')
                    time.sleep(1)
                    ### DOWNLOAD DO ARQUIVO ###
                    ftp = client.open_sftp()
                    ftp.get(FILE_SD, LOCAL_SD)
                    time.sleep(3)
                    print(colored('\rDOWNLOAD: OK ', 'green'))
                    ftp.close()
                    print()
                else:
                    print(colored('PASTA SYSTEM NO SDCARD NAO LOCALIZADA\n-> SDCARD COM DEFEITO!!!\n', 'red'))
                    time.sleep(1)
                    SD_PLAN = 'SDCARD COM DEFEITO'
                    LOCALIZADOR = open(g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\'+g_OS+' - '+g_NS+'.txt', 'a')
                    LOCALIZADOR.write('\n-> SDCARD COM DEFEITO')
                    LOCALIZADOR.close()          
            
        ### INICIANDO BKP DA UD EM OPT (COMPULAB) ###
        print(colored("---> VERIFICANDO A PASTA UD NO COMPULAB <---", 'blue', attrs=['bold']))
        stdin, stdout, stderr = client.exec_command('find /opt/v750/system/ud')
        time.sleep(2)
        UD = stdout.read().decode('utf8')
        print(UD[0:19])
            
        if UD[0:19] == "/opt/v750/system/ud":
                print(colored("   !!! PASTA UD - OK !!!\n", 'green'))
                print(colored("---> INICIANDO BKP DA PASTA UD <---", 'blue', attrs=['bold']))
                stdin, stdout, stderr = client.exec_command("cd /opt/v750/system/ud \n""tar --exclude='photos' -cf Ud.tar.gz *")
                time.sleep(3)
                print(colored('Arquivo TAR do COMPULAB:', attrs=['bold']))
                ssh_command(client, "cd /opt/v750/system/ud \n""ls -lh Ud*")
                print('DOWNLOAD: ...', end='')
                time.sleep(1)
                ### DOWNLOAD DO ARQUIVO ###
                ftp = client.open_sftp()
                ftp.get(FILE_UD, LOCAL_UD)
                time.sleep(4)
                print(colored('\rDOWNLOAD: OK ', 'green'))
                ftp.close()
                print()
        else:
                print(colored('PASTA UD NO COMPULAB NAO LOCALIZADA:\n-> EQUIPAMENTO COM FALHA!!!', 'red'))
                BKP_PLAN = ' / COMPULAB COM DEFEITO'
                print()
                time.sleep(1)

        ### VERIFICANDO ARQUIVOS PCGAR ###
        print(colored('---> VERIFICANDO ARQUIVOS PCGAR <---', 'blue', attrs=['bold']))
        stdin, stdout, stderr = client.exec_command("cd /opt/v750/system/ud/\n""du -hsb File_002_00000\n")
        KB_00000 = stdout.read().decode('utf8')
        print(KB_00000[:-15]+'File_002_00000')      
        KB = KB_00000.find('File_002_00000') 

        if KB == -1:
                print(colored("FALHA AO LOCALIZAR ARQUIVO File_002_0000\n", 'red'))
                time.sleep(1.5)
                SEQUENCIAL = ' '

        elif KB >= 0:
                if (int(KB_00000[:-15])) <= 1000:
                    print(colored('    --> File_002_00000 MENOR QUE 1000Kb <--\n', 'red'))
                    time.sleep(1.5)
                    SEQUENCIAL = ' '

                else:
                    print(colored('\n    --> INICIANDO RESGATE DO ARQUIVO File_002_00000 <--', 'blue', attrs=['bold']))
                    stdin, stdout, stderr = client.exec_command("cd /opt/v750/system/ud\n"f"/root/app/{EXECUTAVELX} -u")
                    SPTRANS_U = stdout.read().decode('utf8')
                    print(SPTRANS_U[2092:2685]+'\n')
                    time.sleep(2)
                    SEQUENCIAL = int(re.search(r'\d+', SPTRANS_U[2365:2410]).group())

                    print(colored(f'SEQUENCIAL:', attrs=['bold']), end=" ")
                    print(colored(f'{SEQUENCIAL}', 'green'))
                    time.sleep(1)

                    if (len(str(SEQUENCIAL))) == 5:
                        pass
                    elif (len(str(SEQUENCIAL))) == 4:
                        SEQUENCIAL = '0'+(str(SEQUENCIAL))
                    elif (len(str(SEQUENCIAL))) == 3:
                        SEQUENCIAL = '00'+(str(SEQUENCIAL))
                    elif (len(str(SEQUENCIAL))) == 2:
                        SEQUENCIAL = '000'+(str(SEQUENCIAL))
                    elif (len(str(SEQUENCIAL))) == 1:
                        SEQUENCIAL = '0000'+(str(SEQUENCIAL))
                    time.sleep(1)

                    ### RENOMEANDO FILE_102_00000 ###
                    print(colored("\n    --> RENOMEANDO FILE 102 <--", 'blue', attrs=['bold']))
                    stdin, stdout, stderr = client.exec_command("cd /opt/v750/system/ud\n" f"mv File_102_00000 File_102_{SEQUENCIAL}")
                    print(f'File_112_00000 - Renomeado para:', end=' ')
                    print(colored(f'File_112_{SEQUENCIAL}\n', 'green'))
                    SEQUENCIAL_INFO = SEQUENCIAL

        ### CRIANDO PCGAR'S ###
        print(colored("---> VERIFICANDO E GERANDO DEMAIS ARQUIVOS PCGAR'S <---", 'blue', attrs=['bold']))
        stdin, stdout, stderr = client.exec_command("cd /opt/v750/system/ud\n"f"/root/app/{EXECUTAVELX} -R")
        time.sleep(4)
        stdin, stdout, stderr = client.exec_command("cd /opt/v750/system/ud\n""ls PCGAR_*")
        PCGAR = stdout.read().decode('utf8')
        print(PCGAR[:-1])

        if PCGAR[:5] == 'PCGAR':
                print(colored('SALVANDO OS ARQUIVOS PCGAR EXISTENTES!!!', 'green'))
                ssh_command(client, "cd /opt/v750/system/ud\n""tar -cf PCGAR.tar.gz PCGAR_*")
                print('DOWNLOAD: ...', end='')
                time.sleep(2)
                ### DOWNLOAD DO ARQUIVO ###
                ftp = client.open_sftp()
                ftp.get(FILE_PCGAR, LOCAL_PCGAR)
                time.sleep(3)
                print(colored('\rDOWNLOAD: OK ', 'green'))
                ftp.close()
                print()
                INFO_PCGAR_TG = 'PCGAR'
                SEQUENCIAL_PCGAR = ' '
                NLINHA = len(PCGAR)
                if NLINHA > 60:
                    ULTIMOPCGAR = NLINHA-9
                    SEQUENCIAL_PCGAR = (PCGAR[45:50]+' AO '+PCGAR[ULTIMOPCGAR:-4])
                    SEQUENCIAL_INFO = ' '
                elif NLINHA == 54:
                    SEQUENCIAL_PCGAR = (PCGAR[45:50])
        else:
                print(colored('SEM ARQUIVOS PCGAR PARA RESGATAR!!!\n', 'red'))
                INFO_PCGAR_TG = 'NAO'
                SEQUENCIAL_INFO = ' '
                SEQUENCIAL_PCGAR = ' '																
                time.sleep(2)

    ############################################## CMT-PROJETOS ##################################################
    elif g_SP_CMT2 == 'CMT-PROJETOS':
            
            ######################## INFO PLANILHA PARA NAO FALHAR ########################
            OBS_PLAN = ' '
            INFO_PCGAR_BOM = ' '
            INFO_PCGAR_TG = ' '
            SEQUENCIAL_INFO = ' '
            SEQUENCIAL_BOM = ' '
            SEQUENCIAL_PCGAR = ' '
            SERIAL_PLAN = ' '

            ### VERIFICANDO NUMERO DE SERIE INTERNO ###
            print(colored("---> VERIFICANDO NUMERO DE SERIE <---", 'blue', attrs=['bold']))
            stdin, stdout, stderr = client.exec_command('ls serial.txt')
            SERIAL = stdout.read().decode('utf8')
            SERIAL_SERIALREADER = SERIAL.find("serial.txt")

            if SERIAL_SERIALREADER >= 0:
                    stdin, stdout, stderr = client.exec_command("tail -1 serial.txt\n")
                    SERIAL = stdout.read().decode('utf8')

                    if SERIAL == "":
                        print(colored('Serial do Validador: ARQUIVO TXT SEM INFORMAÇÃO', attrs=['bold']))
                        print('')
                        ### DOWNLOAD DO TXT ###
                        ftp = client.open_sftp()
                        ftp.get(FILE_TXT, PASTA_TXT)
                        ftp.close()
                    elif SERIAL[:5] == g_NS:
                        print(colored('Serial do Validador:', attrs=['bold']), end=' ')
                        print(SERIAL[:5])
                        print(colored('SERIAL COMPATIVEL - OK', 'green'))
                        print('')
                        ### DOWNLOAD DO TXT ###
                        ftp = client.open_sftp()
                        ftp.get(FILE_TXT, PASTA_TXT)
                        ftp.close()
                    elif '0'+SERIAL[:4] == g_NS:
                        print(colored('Serial do Validador:', attrs=['bold']), end=' ')
                        print(SERIAL[:4])
                        print(colored('SERIAL COMPATIVEL - OK', 'green'))
                        print('')
                        ### DOWNLOAD DO TXT ###
                        ftp = client.open_sftp()
                        ftp.get(FILE_TXT, PASTA_TXT)
                        ftp.close()
                    elif SERIAL[:4] == g_NS:
                        print(colored('Serial do Validador:', attrs=['bold']), end=' ')
                        print(SERIAL[:4])
                        print(colored('SERIAL COMPATIVEL - OK', 'green'))
                        print('')
                        ### DOWNLOAD DO TXT ###
                        ftp = client.open_sftp()
                        ftp.get(FILE_TXT, PASTA_TXT)
                        ftp.close()
                    elif SERIAL[:3] == g_NS:
                        print(colored('Serial do Validador:', attrs=['bold']), end=' ')
                        print(SERIAL[:3])
                        print(colored('SERIAL COMPATIVEL - OK', 'green'))
                        print('')
                        ### DOWNLOAD DO TXT ###
                        ftp = client.open_sftp()
                        ftp.get(FILE_TXT, PASTA_TXT)
                        ftp.close()
                    elif SERIAL[:2] == g_NS:
                        print(colored('Serial do Validador:', attrs=['bold']), end=' ')
                        print(SERIAL[:2])
                        print(colored('SERIAL COMPATIVEL - OK', 'green'))
                        print('')
                        ### DOWNLOAD DO TXT ###
                        ftp = client.open_sftp()
                        ftp.get(FILE_TXT, PASTA_TXT)
                        ftp.close()
                    elif SERIAL[:1] == g_NS:
                        print(colored('Serial do Validador:', attrs=['bold']), end=' ')
                        print(SERIAL[:1])
                        print(colored('SERIAL COMPATIVEL - OK', 'green'))
                        print('')
                        ### DOWNLOAD DO TXT ###
                        ftp = client.open_sftp()
                        ftp.get(FILE_TXT, PASTA_TXT)
                        ftp.close()
                    else:
                        if '0' in SERIAL:
                            print(colored('DIVERGENTE DO BACKSHELL:', 'red', attrs=['bold']))
                            print(colored('Serial OS:', attrs=['bold']), end=' ')
                            print(g_NS)
                            print(colored('Serial Validador:', attrs=['bold']), end=' ')
                            print(SERIAL[:5])
                            print('')
                            ### DOWNLOAD DO TXT ###
                            ftp = client.open_sftp()
                            ftp.get(FILE_TXT, PASTA_TXT)
                            ftp.close()
                            SERIAL_PLAN = ('NS: '+SERIAL[:5])

            else:
                    stdin, stdout, stderr = client.exec_command('ls serialreader.txt')
                    SERIAL2 = stdout.read().decode('utf8')
                    SERIAL_READER = SERIAL2.find("serialreader.txt")
                    
                    if SERIAL_READER >= 0:
                        stdin, stdout, stderr = client.exec_command("tail -1 serialreader.txt\n")
                        SERIALREAD = stdout.read().decode('utf8')

                        if SERIALREAD == "":
                            print(colored('Serial do Validador: ARQUIVO TXT SEM INFORMAÇÃO', attrs=['bold']))
                            print('')
                            ### DOWNLOAD DO TXT ###
                            ftp = client.open_sftp()
                            ftp.get(FILE_TXT_CMT, PASTA_TXT_CMT)
                            ftp.close()
                        elif SERIALREAD[:5] == g_NS:
                            print(colored('Serial do Validador:', attrs=['bold']), end=' ')
                            print(SERIALREAD[:5])
                            print(colored('SERIAL COMPATIVEL - OK', 'green'))
                            print('')
                            ### DOWNLOAD DO TXT ###
                            ftp = client.open_sftp()
                            ftp.get(FILE_TXT_CMT, PASTA_TXT_CMT)
                            ftp.close()
                        elif '0'+SERIALREAD[:4] == g_NS:
                            print(colored('Serial do Validador:', attrs=['bold']), end=' ')
                            print(SERIALREAD[:4])
                            print(colored('SERIAL COMPATIVEL - OK', 'green'))
                            print('')
                            ### DOWNLOAD DO TXT ###
                            ftp = client.open_sftp()
                            ftp.get(FILE_TXT_CMT, PASTA_TXT_CMT)
                            ftp.close()
                        elif SERIALREAD[:4] == g_NS:
                            print(colored('Serial do Validador:', attrs=['bold']), end=' ')
                            print(SERIALREAD[:4])
                            print(colored('SERIAL COMPATIVEL - OK', 'green'))
                            print('')
                            ### DOWNLOAD DO TXT ###
                            ftp = client.open_sftp()
                            ftp.get(FILE_TXT_CMT, PASTA_TXT_CMT)
                            ftp.close()
                        elif SERIALREAD[:3] == g_NS:
                            print(colored('Serial do Validador:', attrs=['bold']), end=' ')
                            print(SERIALREAD[:3])
                            print(colored('SERIAL COMPATIVEL - OK', 'green'))
                            print('')
                            ### DOWNLOAD DO TXT ###
                            ftp = client.open_sftp()
                            ftp.get(FILE_TXT_CMT, PASTA_TXT_CMT)
                            ftp.close()
                        elif SERIALREAD[:2] == g_NS:
                            print(colored('Serial do Validador:', attrs=['bold']), end=' ')
                            print(SERIALREAD[:2])
                            print(colored('SERIAL COMPATIVEL - OK', 'green'))
                            print('')
                            ### DOWNLOAD DO TXT ###
                            ftp = client.open_sftp()
                            ftp.get(FILE_TXT_CMT, PASTA_TXT_CMT)
                            ftp.close()
                        elif SERIALREAD[:1] == g_NS:
                            print(colored('Serial do Validador:', attrs=['bold']), end=' ')
                            print(SERIALREAD[:1])
                            print(colored('SERIAL COMPATIVEL - OK', 'green'))
                            print('')
                            ### DOWNLOAD DO TXT ###
                            ftp = client.open_sftp()
                            ftp.get(FILE_TXT_CMT, PASTA_TXT_CMT)
                            ftp.close()
                        else:
                            if '0' in SERIALREAD:
                                print(colored('DIVERGENTE DO BACKSHELL:', 'red', attrs=['bold']))
                                print(colored('Serial OS:', attrs=['bold']), end=' ')
                                print(g_NS)
                                print(colored('Serial Validador:', attrs=['bold']), end=' ')
                                print(SERIALREAD[:5])
                                print('')
                                ### DOWNLOAD DO TXT ###
                                ftp = client.open_sftp()
                                ftp.get(FILE_TXT_CMT, PASTA_TXT_CMT)
                                ftp.close()
                                SERIAL_PLAN = ('NS: '+SERIALREAD[:5])
                    else:
                        print(colored('ARQUIVO TXT NAO ENCONTRADO!', 'red', attrs=['bold']))
                        print('')

            ### VERIFICANDO A EXISTENCIA DO SDCARD ###
            print(colored("---> VERIFICANDO SDCARD <---", 'blue', attrs=['bold']))
            time.sleep(1)
            stdin, stdout, stderr = client.exec_command("fdisk -l /dev/mmcblk0p1")
            SDCARD = stdout.read().decode('utf8')
            SD = SDCARD.find('/dev/mmcblk0p1')
            print(SDCARD[:29], end=' ')
            
            if SD == -1:
                ### SDCARD INOPERANTE OU NAO EXISTE ###
                print(colored('\r   !!! SDCard não detectado !!!\n', 'red'))
                time.sleep(1)
                SD_PLAN = 'SDCARD COM DEFEITO'

                LOCALIZADOR = open(g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\'+g_OS+' - '+g_NS+'.txt', 'a')
                LOCALIZADOR.write('\n-> SDCARD COM DEFEITO')
                LOCALIZADOR.close()
            else:
                ### INICIANDO BKP DA SYSTEM EM SDCARD ###
                print(colored('\n    !!! SDCard - OK !!!\n', 'green'))
                SD_PLAN = ' '
                time.sleep(1)
                print(colored("---> INICIANDO BKP DA PASTA SYSTEM EM SDCARD <---", 'blue', attrs=['bold']))
                stdin, stdout, stderr = client.exec_command('find /media/mmcblk0p1/opt/v750/system/ud')
                time.sleep(2)
                try:
                    VERFSD = stdout.read().decode('utf8')
                except:
                    VERFSD = 'SDCARD COM DEFEITO'
                
                if VERFSD[0:35] == "/media/mmcblk0p1/opt/v750/system/ud":
                    stdin, stdout, stderr = client.exec_command('cd /media/mmcblk0p1/opt/v750/system/ud/backup/ \n''ls')
                    VERFSD2 = stdout.read().decode('utf8')
                    if VERFSD2 == "":
                        print(colored('AVISO: Pasta BACKUP do SDCARD está vazia!', 'yellow', attrs=['bold']))
                    else:
                        pass
                    
                    stdin, stdout, stderr = client.exec_command("cd /media/mmcblk0p1/opt/v750/system/ud \n""tar --exclude='photos' -cf System.tar.gz *")
                    time.sleep(8)
                    print(colored('Arquivo TAR do SDCARD:', attrs=['bold']))
                    ssh_command(client, "cd /media/mmcblk0p1/opt/v750/system/ud \n""ls -lh System*")
                    print('DOWNLOAD: ...', end='')
                    time.sleep(1)
                    ### DOWNLOAD DO ARQUIVO ###
                    ftp = client.open_sftp()
                    ftp.get(FILE_SD, LOCAL_SD)
                    time.sleep(3)
                    print(colored('\rDOWNLOAD: OK ', 'green'))
                    ftp.close()
                    print()
                else:
                    print(colored('PASTA SYSTEM NO SDCARD NAO LOCALIZADA\n-> SDCARD COM DEFEITO!!!\n', 'red'))
                    time.sleep(1)
                    SD_PLAN = 'SDCARD COM DEFEITO'
                    LOCALIZADOR = open(g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\'+g_OS+' - '+g_NS+'.txt', 'a')
                    LOCALIZADOR.write('\n-> SDCARD COM DEFEITO')
                    LOCALIZADOR.close()
            
            ### INICIANDO BKP DA UD EM OPT (COMPULAB) ###
            print(colored("---> VERIFICANDO A PASTA UD NO COMPULAB <---", 'blue', attrs=['bold']))
            #ssh_command(client, '/opt/v750/system/ud/')
            stdin, stdout, stderr = client.exec_command('find /opt/v750/system/ud')
            time.sleep(2)
            UD = stdout.read().decode('utf8')
            print(UD[0:19]+'\n')
            
            if UD[0:19] == "/opt/v750/system/ud":
                print(colored("---> INICIANDO BKP DA PASTA UD <---", 'blue', attrs=['bold']))
                stdin, stdout, stderr = client.exec_command("cd /opt/v750/system/ud \n""tar -cf Ud.tar.gz File_*")
                time.sleep(2)
                print(colored('Arquivo TAR do COMPULAB:', attrs=['bold']))
                ssh_command(client, "cd /opt/v750/system/ud \n""ls -lh Ud*")
                print('DOWNLOAD: ...', end='')
                time.sleep(2)
                ### DOWNLOAD DO ARQUIVO ###
                ftp = client.open_sftp()
                ftp.get(FILE_UD, LOCAL_UD)
                time.sleep(3)
                print(colored('\rDOWNLOAD: OK ', 'green'))
                ftp.close()
                print()
                print()
            else:
                print(colored('PASTA UD NO COMPULAB NAO LOCALIZADA:\n-> EQUIPAMENTO COM FALHA!!!', 'red'))
                print()
                time.sleep(1)

    ### VERIFICANDO ARQUIVOS TG ###
            print(colored('   ---> VERIFICANDO FILE 002 - TG <---', 'blue', attrs=['bold']))
            stdin, stdout, stderr = client.exec_command("cd /opt/v750/system/ud/\n""du -hsb File_002_00000")
            KB_00000 = stdout.read().decode('utf8')
            print(KB_00000)
            print(colored("   --> VERIFICANDO E GERANDO ARQUIVOS TG'S <--", 'blue', attrs=['bold']))
            stdin, stdout, stderr = client.exec_command("cd /opt/v750/system/ud\n"f"/root/app/{EXECUTAVELX} -R")
            time.sleep(2)
            stdin, stdout, stderr = client.exec_command("cd /opt/v750/system/ud\n""ls TG*")
            TG = stdout.read().decode('utf8')
            print(TG[:-1])
                    
            if (TG[:2]) == 'TG':
                print(colored("SALVANDO OS ARQUIVOS TG EXISTENTES!!!", 'green'))
                ssh_command(client, "cd /opt/v750/system/ud\n""tar -cf TG.tar.gz TG*")
                print('DOWNLOAD: ...', end='')
                time.sleep(2)
                ### DOWNLOAD DO ARQUIVO ###
                ftp = client.open_sftp()
                ftp.get(FILE_TG, LOCAL_TG)
                time.sleep(3)
                print(colored('\rDOWNLOAD: OK ', 'green'))
                ftp.close()
                INFO_PCGAR_TG = 'TG'
                SEQUENCIAL_INFO = (TG[18:23])
                NLINHA = len(TG)
                if NLINHA > 55:
                    ULTIMO_TG = NLINHA-35
                    SEQUENCIAL_INFO = (TG[18:23]+' AO '+TG[ULTIMO_TG:-30])
                #elif NLINHA == 52:
                #    SEQUENCIAL_INFO = (TG[18:23])
            else:
                print(colored('SEM ARQUIVOS TG PARA RESGATAR!!!\n', 'red'))
                INFO_PCGAR_TG = 'NAO'
                time.sleep(2)


    ############################################## TOP - BOM #####################################################
    elif g_SP_CMT2 == 'TOP-BOM':

            ######################## INFO PLANILHA PARA NAO FALHAR ########################
            OBS_PLAN = ' '
            INFO_PCGAR_TG = ' '
            INFO_PCGAR_BOM = ' '
            SEQUENCIAL_INFO = ' '
            SEQUENCIAL_PCGAR = ' '
            SERIAL_PLAN = ' '

            ### VERIFICANDO NUMERO DE SERIE INTERNO ###
            print(colored("---> VERIFICANDO NUMERO DE SERIE <---", 'blue', attrs=['bold']))
            stdin, stdout, stderr = client.exec_command('ls serial.txt')
            SERIAL = stdout.read().decode('utf8')
            SERIAL_SERIALREADER = SERIAL.find("serial.txt")

            if SERIAL_SERIALREADER >= 0:
                stdin, stdout, stderr = client.exec_command("tail -1 serial.txt\n")
                SERIAL = stdout.read().decode('utf8')

                if SERIAL[:5] == g_NS:
                    print(colored('Serial do Validador:', attrs=['bold']), end=' ')
                    print(SERIAL[:5])
                    print(colored('SERIAL COMPATIVEL - OK', 'green'))
                    print('')
                    ### DOWNLOAD DO TXT ###
                    ftp = client.open_sftp()
                    ftp.get(FILE_TXT, PASTA_TXT)
                    ftp.close()
                elif '0'+SERIAL[:4] == g_NS:
                    print(colored('Serial do Validador:', attrs=['bold']), end=' ')
                    print(SERIAL[:4])
                    print(colored('SERIAL COMPATIVEL - OK', 'green'))
                    print('')
                    ### DOWNLOAD DO TXT ###
                    ftp = client.open_sftp()
                    ftp.get(FILE_TXT, PASTA_TXT)
                    ftp.close()
                elif SERIAL[:4] == g_NS:
                    print(colored('Serial do Validador:', attrs=['bold']), end=' ')
                    print(SERIAL[:4])
                    print(colored('SERIAL COMPATIVEL - OK', 'green'))
                    print('')
                    ### DOWNLOAD DO TXT ###
                    ftp = client.open_sftp()
                    ftp.get(FILE_TXT, PASTA_TXT)
                    ftp.close()
                elif SERIAL[:3] == g_NS:
                    print(colored('Serial do Validador:', attrs=['bold']), end=' ')
                    print(SERIAL[:3])
                    print(colored('SERIAL COMPATIVEL - OK', 'green'))
                    print('')
                    ### DOWNLOAD DO TXT ###
                    ftp = client.open_sftp()
                    ftp.get(FILE_TXT, PASTA_TXT)
                    ftp.close()
                elif SERIAL[:2] == g_NS:
                    print(colored('Serial do Validador:', attrs=['bold']), end=' ')
                    print(SERIAL[:2])
                    print(colored('SERIAL COMPATIVEL - OK', 'green'))
                    print('')
                    ### DOWNLOAD DO TXT ###
                    ftp = client.open_sftp()
                    ftp.get(FILE_TXT, PASTA_TXT)
                    ftp.close()
                elif SERIAL[:1] == g_NS:
                    print(colored('Serial do Validador:', attrs=['bold']), end=' ')
                    print(SERIAL[:1])
                    print(colored('SERIAL COMPATIVEL - OK', 'green'))
                    print('')
                    ### DOWNLOAD DO TXT ###
                    ftp = client.open_sftp()
                    ftp.get(FILE_TXT, PASTA_TXT)
                    ftp.close()
                else:
                    if '0' in SERIAL:
                        print(colored('DIVERGENTE DO BACKSHELL:', 'red', attrs=['bold']))
                        print(colored('Serial OS:', attrs=['bold']), end=' ')
                        print(g_NS)
                        print(colored('Serial Validador:', attrs=['bold']), end=' ')
                        print(SERIAL[:5])
                        print('')
                        ### DOWNLOAD DO TXT ###
                        ftp = client.open_sftp()
                        ftp.get(FILE_TXT, PASTA_TXT)
                        ftp.close()
                        SERIAL_PLAN = ('NS: '+SERIAL[:5])

            else:
                stdin, stdout, stderr = client.exec_command('ls serialreader.txt')
                SERIAL2 = stdout.read().decode('utf8')
                SERIAL_READER = SERIAL2.find("serialreader.txt")
                
                if SERIAL_READER >= 0:
                    stdin, stdout, stderr = client.exec_command("tail -1 serialreader.txt\n")
                    SERIALREAD = stdout.read().decode('utf8')

                    if SERIALREAD[:5] == g_NS:
                        print(colored('Serial do Validador:', attrs=['bold']), end=' ')
                        print(SERIALREAD[:5])
                        print(colored('SERIAL COMPATIVEL - OK', 'green'))
                        print('')
                        ### DOWNLOAD DO TXT ###
                        ftp = client.open_sftp()
                        ftp.get(FILE_TXT_CMT, PASTA_TXT_CMT)
                        ftp.close()
                    elif '0'+SERIALREAD[:4] == g_NS:
                        print(colored('Serial do Validador:', attrs=['bold']), end=' ')
                        print(SERIALREAD[:4])
                        print(colored('SERIAL COMPATIVEL - OK', 'green'))
                        print('')
                        ### DOWNLOAD DO TXT ###
                        ftp = client.open_sftp()
                        ftp.get(FILE_TXT_CMT, PASTA_TXT_CMT)
                        ftp.close()
                    elif SERIALREAD[:4] == g_NS:
                        print(colored('Serial do Validador:', attrs=['bold']), end=' ')
                        print(SERIALREAD[:4])
                        print(colored('SERIAL COMPATIVEL - OK', 'green'))
                        print('')
                        ### DOWNLOAD DO TXT ###
                        ftp = client.open_sftp()
                        ftp.get(FILE_TXT_CMT, PASTA_TXT_CMT)
                        ftp.close()
                    elif SERIALREAD[:3] == g_NS:
                        print(colored('Serial do Validador:', attrs=['bold']), end=' ')
                        print(SERIALREAD[:3])
                        print(colored('SERIAL COMPATIVEL - OK', 'green'))
                        print('')
                        ### DOWNLOAD DO TXT ###
                        ftp = client.open_sftp()
                        ftp.get(FILE_TXT_CMT, PASTA_TXT_CMT)
                        ftp.close()
                    elif SERIALREAD[:2] == g_NS:
                        print(colored('Serial do Validador:', attrs=['bold']), end=' ')
                        print(SERIALREAD[:2])
                        print(colored('SERIAL COMPATIVEL - OK', 'green'))
                        print('')
                        ### DOWNLOAD DO TXT ###
                        ftp = client.open_sftp()
                        ftp.get(FILE_TXT_CMT, PASTA_TXT_CMT)
                        ftp.close()
                    elif SERIALREAD[:1] == g_NS:
                        print(colored('Serial do Validador:', attrs=['bold']), end=' ')
                        print(SERIALREAD[:1])
                        print(colored('SERIAL COMPATIVEL - OK', 'green'))
                        print('')
                        ### DOWNLOAD DO TXT ###
                        ftp = client.open_sftp()
                        ftp.get(FILE_TXT_CMT, PASTA_TXT_CMT)
                        ftp.close()
                    else:
                        if '0' in SERIALREAD:
                            print(colored('DIVERGENTE DO BACKSHELL:', 'red', attrs=['bold']))
                            print(colored('Serial OS:', attrs=['bold']), end=' ')
                            print(g_NS)
                            print(colored('Serial Validador:', attrs=['bold']), end=' ')
                            print(SERIALREAD[:5])
                            print('')
                            ### DOWNLOAD DO TXT ###
                            ftp = client.open_sftp()
                            ftp.get(FILE_TXT_CMT, PASTA_TXT_CMT)
                            ftp.close()
                            SERIAL_PLAN = ('NS: '+SERIALREAD[:5])
                else:
                    print(colored('ARQUIVO TXT NAO ENCONTRADO!', 'red', attrs=['bold']))
                    print('')
                    SERIAL_PLAN = ' '

            ### VERIFICANDO A EXISTENCIA DO SDCARD ###
            print(colored("---> VERIFICANDO SDCARD <---", 'blue', attrs=['bold']))
            time.sleep(1)
            stdin, stdout, stderr = client.exec_command("fdisk -l /dev/mmcblk0p1")
            SDCARD = stdout.read().decode('utf8')
            SD = SDCARD.find('/dev/mmcblk0p1')
            print(SDCARD[:29], end=' ')
            
            if SD == -1:
                ### SDCARD INOPERANTE OU NAO EXISTE ###
                print(colored('\r   !!! SDCard não detectado !!!\n', 'red'))
                time.sleep(1)
                SD_PLAN = 'SDCARD COM DEFEITO'
                ### GRAVANDO INFORMAÇÃO NO TXT ###
                LOCALIZADOR = open(g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\'+g_OS+' - '+g_NS+'.txt', 'a')
                LOCALIZADOR.write('\n-> SDCARD COM DEFEITO')
                LOCALIZADOR.close()
            else:
                ### INICIANDO BKP DA SYSTEM EM SDCARD ###
                print(colored('\n    !!! SDCard - OK !!!\n', 'green'))
                SD_PLAN = ' '
                time.sleep(1)
                print(colored("---> INICIANDO BKP DA PASTA SYSTEM EM SDCARD <---", 'blue', attrs=['bold']))
                stdin, stdout, stderr = client.exec_command('fdisk -l /dev/mmcblk0p1')
                time.sleep(2)
                try:
                    VERFSD = stdout.read().decode('utf8')
                except:
                    VERFSD = 'SDCARD COM DEFEITO'
                
                if VERFSD[0:16] == "/media/mmcblk0p1":
                    stdin, stdout, stderr = client.exec_command('cd /media/mmcblk0p1/opt/v750/system/ud/backup/ \n''ls')
                    VERFSD2 = stdout.read().decode('utf8')
                    if VERFSD2 == "":
                        print(colored('AVISO: Pasta BACKUP do SDCARD está vazia!', 'yellow', attrs=['bold']))
                    else:
                        pass
                    stdin, stdout, stderr = client.exec_command('cd /media/mmcblk0p1/001/opt/v750/system/ud/ \n''ls')
                    VERFSD2 = stdout.read().decode('utf8')
                    if VERFSD2 == "":
                        print(colored('AVISO: Pasta BOM do SDCARD está vazia!', 'yellow', attrs=['bold']))
                    else:
                        pass
                    
                    stdin, stdout, stderr = client.exec_command("cd /media/mmcblk0p1/\n""tar --exclude='photos' -cf mmcblk0p1.tar.gz *")
                    time.sleep(8)
                    print(colored('Arquivo TAR do SDCARD:', attrs=['bold']))
                    ssh_command(client, "cd /media/mmcblk0p1/\n""ls -lh mmcblk0p1*")
                    print('DOWNLOAD: ...', end='')
                    time.sleep(1)
                    ### DOWNLOAD DO ARQUIVO ###
                    ftp = client.open_sftp()
                    ftp.get(FILE_SD_TOP, LOCAL_SD_TOP)
                    time.sleep(3)
                    print(colored('\rDOWNLOAD: OK ', 'green'))
                    ftp.close()
                    print()
                else:
                    print(colored('PASTA SYSTEM NO SDCARD NAO LOCALIZADA\n-> SDCARD COM DEFEITO!!!\n', 'red'))
                    time.sleep(1)
                    SD_PLAN = 'SDCARD COM DEFEITO'
                    LOCALIZADOR = open(g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\'+g_OS+' - '+g_NS+'.txt', 'a')
                    LOCALIZADOR.write('\n-> SDCARD COM DEFEITO')
                    LOCALIZADOR.close()
            
            ########## INICIANDO BKP DA UD EM OPT (TOP - COMPULAB) ##########
            time.sleep(3)
            print(colored("---> VERIFICANDO A PASTA TOP NO COMPULAB <---", 'blue', attrs=['bold']))
            #ssh_command(client, '/opt/v750/system/ud/')
            stdin, stdout, stderr = client.exec_command('find /opt/v750/system/ud')
            time.sleep(2)
            UD = stdout.read().decode('utf8')
            print(UD[0:19]+'\n')
            
            if UD[0:19] == "/opt/v750/system/ud":
                print(colored("\n---> INICIANDO BKP DA PASTA TOP <---", 'blue', attrs=['bold']))
                stdin, stdout, stderr = client.exec_command("cd /opt/v750/system/ud \n""tar -cf TOP.tar.gz *")
                time.sleep(2)
                print(colored('Arquivo TAR do COMPULAB (TOP):', attrs=['bold']))
                ssh_command(client, "cd /opt/v750/system/ud \n""ls -lh TOP*")
                print('DOWNLOAD: ...', end='')
                time.sleep(2)
                ### DOWNLOAD DO ARQUIVO ###
                ftp = client.open_sftp()
                ftp.get(FILE_TOP, LOCAL_TOP)
                time.sleep(3)
                print(colored('\rDOWNLOAD: OK ', 'green'))
                ftp.close()
                print()
            else:
                print(colored('PASTA TOP NO COMPULAB NAO LOCALIZADA:\n-> EQUIPAMENTO COM FALHA!', 'red'))
                print()
                time.sleep(2)

            ########## INICIANDO BKP DA UD EM 001 (BOM - COMPULAB) ##########
            time.sleep(3)
            print(colored("---> VERIFICANDO A PASTA BOM NO COMPULAB <---", 'blue', attrs=['bold']))
            #ssh_command(client, '/opt/v750/system/ud/')
            stdin, stdout, stderr = client.exec_command('find /001/opt/v750/system/ud')
            time.sleep(2)
            UD = stdout.read().decode('utf8')
            print(UD[0:23]+'\n')
            
            if UD[0:23] == "/001/opt/v750/system/ud":
                print(colored("\n---> INICIANDO BKP DA PASTA BOM <---", 'blue', attrs=['bold']))
                stdin, stdout, stderr = client.exec_command("cd /001/opt/v750/system/ud\n""tar -cf BOM.tar.gz *")
                time.sleep(2)
                print(colored('Arquivo TAR do BOM:', attrs=['bold']))
                ssh_command(client, "cd /001/opt/v750/system/ud\n""ls -lh BOM*")
                print('DOWNLOAD: ...', end='')
                time.sleep(2)
                ### DOWNLOAD DO ARQUIVO ###
                ftp = client.open_sftp()
                ftp.get(FILE_BOM, LOCAL_BOM)
                time.sleep(3)
                print(colored('\rDOWNLOAD: OK ', 'green'))
                ftp.close()
                print()
            else:
                print(colored('PASTA BOM NO COMPULAB NAO LOCALIZADA:\n-> EQUIPAMENTO COM FALHA!', 'red'))
                print()
                time.sleep(2)

            ### VERIFICANDO ARQUIVOS TG-TOP ###
            print(colored('\n\n   ---> VERIFICANDO FILE 002 - TG (TOP) <---', 'blue', attrs=['bold']))
            stdin, stdout, stderr = client.exec_command("cd /opt/v750/system/ud/\n""du -hsb File_002_*\n")
            KB_00000 = stdout.read().decode('utf8')
            print(KB_00000)
            print(colored("    --> VERIFICANDO E GERANDO ARQUIVOS TG'S TOP <--", 'blue', attrs=['bold']))
            stdin, stdout, stderr = client.exec_command("cd /opt/v750/system/ud\n"f"/root/app/{EXECUTAVELX} -R")
            time.sleep(2)
            stdin, stdout, stderr = client.exec_command("cd /opt/v750/system/ud\n""ls TG*")
            TG = stdout.read().decode('utf8')
            print(TG[:-1])
                    
            if (TG[:2]) == 'TG':
                print(colored("\nSALVANDO OS ARQUIVOS TG-TOP EXISTENTES!!!", 'green'))
                ssh_command(client, "cd /opt/v750/system/ud\n""tar -cf TG.tar.gz TG*")
                print('DOWNLOAD: ...', end='')
                time.sleep(2)
                ### DOWNLOAD DO ARQUIVO ###
                ftp = client.open_sftp()
                ftp.get(FILE_TG, LOCAL_TG)
                time.sleep(3)
                print(colored('\rDOWNLOAD: OK ', 'green'))
                ftp.close()
                INFO_PCGAR_TG = 'TOP'
                SEQUENCIAL_INFO = (TG[18:23])
            else:
                print(colored('SEM ARQUIVOS TG-TOP PARA RESGATAR!!!\n', 'red'))
                INFO_PCGAR_TG = 'NAO'
                SEQUENCIAL_INFO = ' '
                time.sleep(2)

            ### VERIFICANDO ARQUIVOS TG-BOM ###
            print(colored('\n\n   ---> VERIFICANDO FILE 002 - TG (BOM) <---', 'blue', attrs=['bold']))
            stdin, stdout, stderr = client.exec_command("cd /opt/v750/system/ud/\n""du -hsb File_002_*\n")
            KB_00001 = stdout.read().decode('utf8')
            print(KB_00001)
            print(colored("\n    --> VERIFICANDO E GERANDO ARQUIVOS TG'S BOM <--", 'blue', attrs=['bold']))
            stdin, stdout, stderr = client.exec_command("cd /001/opt/v750/system/ud\n"f"/root/app/{EXECUTAVELX} -R")
            time.sleep(2)
            stdin, stdout, stderr = client.exec_command("cd /001/opt/v750/system/ud\n""ls TG*")
            TG_BOM = stdout.read().decode('utf8')
            print(TG_BOM[:-1])
                    
            if (TG_BOM[:2]) == 'TG':
                print(colored("\nSALVANDO OS ARQUIVOS TG-BOM EXISTENTES!!!", 'green'))
                ssh_command(client, "cd /001/opt/v750/system/ud\n""tar -cf TG-BOM.tar.gz TG*")
                print('DOWNLOAD: ...', end='')
                time.sleep(2)
                ### DOWNLOAD DO ARQUIVO ###
                ftp = client.open_sftp()
                ftp.get(FILE_TG_BOM, LOCAL_TG_BOM)
                time.sleep(5)
                print(colored('\rDOWNLOAD: OK ', 'green'))
                ftp.close()
                SEQUENCIAL_BOM = ' / '+(TG_BOM[18:23])
                INFO_PCGAR_BOM = ' / BOM'
            else:
                print(colored('SEM ARQUIVOS TG-BOM PARA RESGATAR!!!\n', 'red'))
                SEQUENCIAL_BOM = ' '
                time.sleep(2)
                
    ##############################################################################################################

    ############################ LIMPANDO RESGISTROS (RAR) DO VALIDADOR ####################################
    print(colored("\n--> LIMPANDO RESGISTROS DO VALIDADOR <--", attrs=['bold']))
    print("LIMPANDO...", end='')
    try:
        stdin, stdout, stderr = client.exec_command(f"rm {FILE_UD}")
        time.sleep(0.5)
        stdin, stdout, stderr = client.exec_command(f"rm {FILE_SD}")
        time.sleep(0.5)
        stdin, stdout, stderr = client.exec_command(f"rm {FILE_PCGAR}")
        time.sleep(0.5)
        stdin, stdout, stderr = client.exec_command(f"rm {FILE_TG}")
        time.sleep(0.5)
        stdin, stdout, stderr = client.exec_command(f"rm {FILE_TG_BOM}")
        time.sleep(0.5)
        stdin, stdout, stderr = client.exec_command(f"rm {FILE_BOM}")
        time.sleep(0.5)
        stdin, stdout, stderr = client.exec_command(f"rm {FILE_TOP}")
        time.sleep(0.5)
        stdin, stdout, stderr = client.exec_command(f"rm {FILE_SD_TOP}")
        time.sleep(0.5)
        print(colored("\rLIMPEZA CONCLUÍDA !!!", 'green'))    
    except:
        print(colored("\rLIMPEZA CONCLUÍDA !!!", 'green'))
    ##############################################################################################################

    ####################################### REATIVANDO APLICAÇÃO #################################################
    stdin, stdout, stderr = client.exec_command(f"/root/app/{EXECUTAVELX}\n")
    
    ############################################# ENCERRANDO CONEXAO #############################################
    print(colored("\n\n     --> FECHANDO CONEXAO SSH. <--", 'yellow'))
    client.close()
    
    ############################ ADICONAR INFORMACOES NO ARQUIVO RESGATE XLSX ####################################
        
    RESGATE = load_workbook(g_C_D+f'\\BKP\\RESGATE-{ANO}.xlsx')
    PLANILHA = RESGATE['INTERVENCAO']
    PLAN_RESG = ''
    
    try:
        ### ADCIONINANDO INFORMACOES NA PLANILHA ###
        RESGATE_INFO = ['F', g_EMPRESA, g_OS, g_NS, INFO_PCGAR_TG + INFO_PCGAR_BOM,SEQUENCIAL_INFO + SEQUENCIAL_PCGAR + SEQUENCIAL_BOM, DATAXLSX, SERIAL_PLAN + BKP_PLAN, SD_PLAN + OBS_PLAN]
        PLANILHA.append(RESGATE_INFO)
        NLINHA = PLANILHA.max_row
        ULTIMA_LINHA = [f'A{NLINHA}', f'B{NLINHA}', f'C{NLINHA}', f'D{NLINHA}', f'E{NLINHA}', f'F{NLINHA}', f'G{NLINHA}', f'H{NLINHA}', f'I{NLINHA}']
        ULTIMA_LINHA2 = [f'B{NLINHA}', f'C{NLINHA}', f'D{NLINHA}', f'E{NLINHA}', f'F{NLINHA}', f'G{NLINHA}', f'H{NLINHA}', f'I{NLINHA}']
        #BORDA = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        for cell in ULTIMA_LINHA2:
            PLANILHA[cell].border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        for cell in ULTIMA_LINHA:
            PLANILHA[cell].alignment = Alignment(horizontal='center')
        RESGATE.save(g_C_D+f'\\BKP\\RESGATE-{ANO}.xlsx')
        print(colored('\n   GRAVADA INFORMAÇÕES NA PLANILHA!!!', 'green'))
        PLAN_RESG = 'OK'
    except:
        print(colored('\nFALHA AO GRAVAR INFORMAÇÕES NA PLANILHA!!!', 'red'))
        print(colored('VERIFIQUE SE A PLANILHA ESTA ABERTA E FECHEA.', 'yellow'))
        PLAN_RESG = 'FALHA AO GRAVAR'
        pass
    
    ############################################# EXTRAÇÃO DOS DOWNLOADS #########################################
    print(colored('\n\n---> INCIANDO EXTRAÇÃO DOS ARQUIVOS DE BKP <---', 'blue',attrs=['bold']))

    ########################## VERIFICANDO SE EXISTE UD PARA EXTRAÇÃO ##########################
    if (os.path.exists(LOCAL_UD)):
        print('\nEXTRAINDO ARQUIVOS UD (COMPULAB)...')
        ###PASTA PARA EXTRAÇÃO###
        # Verificar se este diretorio ja existe
        if os.path.isdir(g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\Ud'):
            print ('Ja existe uma pasta para a UD (Compulab)')
        # Criar a pasta caso nao exista
        else:
            os.mkdir(g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\Ud')
            print(colored('Pasta para a UD (Compulab) criada com sucesso!', 'yellow'))
        
        UD_LOCAL = tarfile.open(LOCAL_UD)
        ### INICIANDO EXTRAÇÃO ###
        try:
            UD_LOCAL.extractall(DESC_UD)
            UD_LOCAL.close()
            print(colored('EXTRAÇÃO COMPLETA - OK', 'green'))
            INFO_UD = 'OK'
        except:
            print(colored("FALHA AO EXTRAIR DA UD (COMPULAB)!", 'red'))
            time.sleep(1)
            INFO_UD = 'FALHA AO EXTRAIR'
            pass
        try:
            os.remove(LOCAL_UD)  
        except:
            pass
    else:
        print(colored("\nNÃO HA ARQUIVOS DA UD (COMPULAB) PARA EXTRAÇÃO...", 'red'))
        time.sleep(1)
        INFO_UD = 'NÃO CONTEM'

    ########################## VERIFICANDO SE EXISTE SDCARD PARA EXTRAÇÃO ##########################
    if(os.path.exists(LOCAL_SD)):
        print('\nEXTRAINDO ARQUIVOS SYSTEM (SDCARD)...')
        ### PASTA PARA EXTRAÇÃO ###
        if os.path.isdir(g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\System'):
            print ('Ja existe uma pasta para a System/SDCARD')
        else:
            os.mkdir(g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\System')
            print(colored("Pasta para a System/SDCARD criada com sucesso!", 'yellow'))
        
        ### INICIANDO EXTRAÇÃO ###
        SDD_LOCAL = tarfile.open(LOCAL_SD)
        try:
            SDD_LOCAL.extractall(DESC_SD)
            SDD_LOCAL.close()
            print(colored('EXTRAÇÃO COMPLETA - OK', 'green'))
            INFO_SD = 'OK'
        except:
            print(colored("FALHA AO EXTRAIR System/SDCARD!", 'red'))
            time.sleep(1)
            INFO_SD = 'FALHA AO EXTRAIR'
            pass
        try:
            os.remove(LOCAL_SD)  
        except:
            pass
    else:           
        print(colored("\nNÃO HA ARQUIVOS DO SDCARD (SPTRANS) PARA EXTRAÇÃO...", 'red'))
        time.sleep(1)
        INFO_SD = 'NÃO CONTÉM'

    ########################## VERIFICANDO SE EXISTE PCGAR PARA EXTRAÇÃO ##########################
    if (os.path.exists(LOCAL_PCGAR)):
        print('\nEXTRAINDO ARQUIVOS PCGAR...')
        PCGAR_LOCAL = tarfile.open(LOCAL_PCGAR)
        
        ### INICIANDO EXTRAÇÃO ###
        try:
            PCGAR_LOCAL.extractall(PASTA_INI)
            PCGAR_LOCAL.close()
            print(colored('EXTRAÇÃO COMPLETA - OK', 'green'))
            INFO_PCGAR = 'PCGAR'
            INFO_TG = ' '
        except:
            print(colored("FALHA AO EXTRAIR PCGAR!", 'red'))
            time.sleep(1)
            INFO_PCGAR = 'FALHA AO EXTRAIR'
            INFO_TG = ' '
            pass
        try:
            os.remove(LOCAL_PCGAR)  
        except:
            pass
    else:
        print(colored("\nNAO HA ARQUIVOS PCGAR PARA EXTRAÇÃO...", 'red'))
        time.sleep(1)
        INFO_PCGAR = 'NÃO CONTÉM'
        INFO_TG = ' '

    ########################## VERIFICANDO SE EXISTE TG PARA EXTRAÇÃO ##########################
    if (os.path.exists(LOCAL_TG)):
        print('\nEXTRAINDO ARQUIVOS TG...')
        TG_LOCAL = tarfile.open(LOCAL_TG)
        
        ### INICIANDO EXTRAÇÃO ###
        try:
            TG_LOCAL.extractall(PASTA_INI)
            TG_LOCAL.close()
            print(colored('EXTRAÇÃO COMPLETA - OK', 'green'))
            INFO_TG = 'TG'
            INFO_PCGAR = ' '
        except:
            print(colored("FALHA AO EXTRAIR TG!", 'red'))
            time.sleep(1)
            INFO_TG = 'FALHA AO EXTRAIR'
            INFO_PCGAR = ' '
            pass
        try:
            os.remove(LOCAL_TG)  
        except:
            pass
    else:
        print(colored("\nNÃO HA ARQUIVOS TG PARA EXTRAÇÃO...", 'red'))
        time.sleep(1)

    ########################## VERIFICANDO SE EXISTE TOP PARA EXTRAÇÃO ##########################
    if (os.path.exists(LOCAL_TOP)):
        print('\nEXTRAINDO ARQUIVOS TOP...')
        ###PASTA PARA EXTRAÇÃO###
        if os.path.isdir(g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\TOP'):
            print ('Ja existe uma pasta TOP')
        else:
            os.mkdir(g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\TOP')
            print(colored('Pasta para o TOP criada com sucesso!', 'yellow'))
        TOP_LOCAL = tarfile.open(LOCAL_TOP)
        
        ### INICIANDO EXTRAÇÃO ###
        try:
            TOP_LOCAL.extractall(DESC_TOP)
            TOP_LOCAL.close()
            print(colored('EXTRAÇÃO COMPLETA - OK', 'green'))
            INFO_TOP = 'OK'
        except:
            print(colored("FALHA AO EXTRAIR PCGAR!", 'red'))
            time.sleep(1)
            INFO_TOP = 'FALHA AO EXTRAIR'
            pass
        try:
            os.remove(LOCAL_TOP)  
        except:
            pass     
    else:
        print(colored("\nNÃO HA ARQUIVOS TOP PARA EXTRAÇÃO...", 'red'))
        INFO_TOP = 'NÃO CONTÉM'
        time.sleep(1)

    ########################## VERIFICANDO SE EXISTE TG-BOM PARA EXTRAÇÃO ##########################
    if (os.path.exists(LOCAL_TG_BOM)):
        print('\nEXTRAINDO ARQUIVOS TG-BOM...')
        ###PASTA PARA EXTRAÇÃO###
        if os.path.isdir(g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\TG-BOM'):
            print ('Ja existe uma pasta para TG-BOM')
        else:
            os.mkdir(g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\TG-BOM')
            print(colored('Pasta para TG-BOM criada com sucesso!', 'yellow'))
        TG_BOM_LOCAL = tarfile.open(LOCAL_TG_BOM)
        
        ### INICIANDO EXTRAÇÃO ###
        try:
            TG_BOM_LOCAL.extractall(DESC_TG_BOM)
            TG_BOM_LOCAL.close()
            print(colored('EXTRAÇÃO COMPLETA - OK', 'green'))
            INFO_BOM = '/ BOM'
        except:
            print(colored("FALHA AO EXTRAIR PCGAR!", 'red'))
            time.sleep(1)
            INFO_BOM = 'FALHA AO EXTRAIR'
            pass
        try:
            os.remove(LOCAL_TG_BOM)  
        except:
            pass       
    else:
        print(colored("\nNÃO HA ARQUIVOS TG-BOM PARA EXTRAÇÃO...", 'red'))
        time.sleep(1)
        INFO_BOM = ' '

    ########################## VERIFICANDO SE EXISTE BOM PARA EXTRAÇÃO ##########################
    if (os.path.exists(LOCAL_BOM)):
        print('\nEXTRAINDO ARQUIVOS BOM (CUMPULAB)...')
        ###PASTA PARA EXTRAÇÃO###
        if os.path.isdir(g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\BOM'):
            print ('Ja existe uma pasta para a BOM (Compulab)')
        else:
            os.mkdir(g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\BOM')
            print(colored('Pasta para a BOM (Compulab) criada com sucesso!\n', 'yellow'))
        BOM_LOCAL = tarfile.open(LOCAL_BOM)
        
        ### INICIANDO EXTRAÇÃO ###
        try:
            BOM_LOCAL.extractall(DESC_BOM)
            BOM_LOCAL.close()
            print(colored('EXTRAÇÃO COMPLETA - OK', 'green'))
            INFO_001 = 'OK'
        except:
            print(colored("FALHA AO EXTRAIR PCGAR!", 'red'))
            time.sleep(1)
            INFO_001 = 'FALHA AO EXTRAIR'
            pass
        try:
            os.remove(LOCAL_BOM)  
        except:
            pass
    else:
        print(colored("\nNÃO HA ARQUIVOS BOM (COMPULAB) PARA EXTRAÇÃO...", 'red'))
        time.sleep(1)
        INFO_001 = 'NÃO CONTÉM'

    ########################## VERIFICANDO SE EXISTE SDCARD - TOP/BOM PARA EXTRAÇÃO ##########################
    if(os.path.exists(LOCAL_SD_TOP)):
        print('\nEXTRAINDO ARQUIVOS SDCARD TOP/BOM...')
        ###PASTA PARA EXTRAÇÃO###
        if os.path.isdir(g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\mmcblk0p1'):
            print ('Ja existe uma pasta para a System/SDCARD')
        else:
            os.mkdir(g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA+'\\mmcblk0p1')
            print(colored('Pasta para a SDCARD TOP/BOM criada com sucesso!', 'yellow'))
        SD_TOP = tarfile.open(LOCAL_SD_TOP)
        
         ### INICIANDO EXTRAÇÃO ###
        try:
            SD_TOP.extractall(DESC_SD_TOP)
            SD_TOP.close()
            print(colored('EXTRAÇÃO COMPLETA - OK', 'green'))
            INFO_SD = 'TOP/BOM'
        except:
            print(colored("FALHA AO EXTRAIR PCGAR!", 'red'))
            time.sleep(1)
            INFO_SD = 'FALHA AO EXTRAIR'
            pass
        try:
            os.remove(LOCAL_SD_TOP)  
        except:
            pass
    else:
        print(colored("\nNÃO HA ARQUIVOS DO SDCARD TOP/BOM PARA EXTRAÇÃO...", 'red'))
        time.sleep(1)

    ######################## COPIAR PARA O SERVIDOR ########################
    print(colored('\n\n--> COPIANDO PARA O SERVIDOR <--', 'blue', attrs=['bold']))
    os.system("net use J: \\\\172.18.10.2\\Assistencia$")
    if os.path.exists(f"J:\#ARQUIVOS#\{g_SP_CMT2}"):
        LOCAL = g_C_D+'\\BKP\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA
        SERV = 'J:\\#ARQUIVOS#\\'+g_SP_CMT2+'\\'+g_EMPRESA+'\\'+g_OS+DATA
        os.listdir(LOCAL)
        print('COPIANDO...', end=' ')
        try:
            shutil.copytree(LOCAL, SERV)
        except:
            pass
        print(colored('\rBKP COM SUCESSO!!! - SERVIDOR (172.18.10.2)', 'green'))
        time.sleep(1)
        os.system('net use J: /delete')
        SERVIDOR = 'OK'
    else:
        print(colored('SEM CONEXÃO COM O SERVIDOR (172.18.10.2)', 'red'))
        SERVIDOR = 'SEM CONEXAO'
        time.sleep(1)
        os.system('net use J: /delete')
        

    ######################## TXT BAIXADO ########################
    if(os.path.exists(TXT)):
        TXT_INFO = 'DOWNLOAD OK!'
        TXT_INFO2 = ' '
    elif(os.path.exists(TXT2)):
        TXT_INFO2 = 'DOWNLOAD OK!'
        TXT_INFO = ' '
    else:
        TXT_INFO = 'SERIAL.TXT NÃO ENCONTRADO!'
        TXT_INFO2 = ' '


    ################################## RELATÓRIO ##################################

    print('\n\n\nFINALIZANDO...')
    print('\nCRIANDO RELATÓRIO...')
    time.sleep(2)
    print()
    print()
    print(colored('===================================================', attrs=['bold']))
    print(colored('                     RELATÓRIO\n', attrs=['bold']))
    print()
    print(colored('  EMPRESA:', 'light_cyan', attrs=['bold']), end=' ')
    print(g_EMPRESA)
    print(colored('  OS:', 'light_cyan', attrs=['bold']), end=' ')
    print(g_OS)
    print(colored('  NS:', 'light_cyan', attrs=['bold']), end=' ')
    print(g_NS)
    print(colored('  TXT SERIAL:', 'light_cyan', attrs=['bold']), end=' ')
    print(TXT_INFO, TXT_INFO2)
    print(colored('  TIPO DE EQUIPMANETO:', 'light_cyan', attrs=['bold']), end=' ')
    print(g_SP_CMT2)
    print(colored('  UD (SPTRANS):', 'light_cyan', attrs=['bold']), end=' ')
    print(INFO_UD)
    print(colored('  SDCARD:', 'light_cyan', attrs=['bold']), end=' ')
    print(INFO_SD)
    print(colored('  PCGAR OU TG:', 'light_cyan', attrs=['bold']), end=' ')
    print(INFO_PCGAR + INFO_TG + INFO_BOM)
    print(colored('  TOP:', 'light_cyan', attrs=['bold']), end=' ')
    print(INFO_TOP)
    print(colored('  BOM:', 'light_cyan', attrs=['bold']), end=' ')
    print(INFO_001)
    print(colored('  BKP PARA SERV. (172.18.10.2):', 'light_cyan', attrs=['bold']), end=' ')
    print(SERVIDOR)
    print(colored('  PLANILHA RESGATE:', 'light_cyan', attrs=['bold']), end=' ')
    print(PLAN_RESG)
    print()     
    print(colored(':               !!!BKP Finalizado!!!              :', 'white', 'on_blue'))
    print(colored('===================================================', attrs=['bold']))
    print()     
    print()     

    ###################### LISTAR PCGAR OU TGS ######################
    try:
        TG_PCGAR = os.listdir(PASTA_INI)
        for f in TG_PCGAR:
            f_path = os.path.join(PASTA_INI,f)
            if(os.path.isfile(f_path)):
                f_size = os.path.getsize(f_path)
                print(f_size, f_path)
    except:
        pass
    try:
        TG_B = os.listdir(DESC_TG_BOM)
        for t in TG_B:
            t_path = os.path.join(DESC_TG_BOM,t)
            if(os.path.isfile(t_path)):
                t_size = os.path.getsize(t_path)
                print(t_size, t_path)
    except:
        pass

    ###################### ABRIR A PASTA ##########################
    os.system(f"explorer.exe "+PASTA_INI)

    ###################### REINICIAR OU NAO ######################
    print(colored("\n\nREINICIAR?", 'yellow'))
    REINICAR_FINAL = input('Digite S para SIM ou ENTER para Fechar...\n').upper()
    if REINICAR_FINAL == 'S':
        print('REINICIANDO...')
        os.system('python RESGATE.py')
    elif REINICAR_FINAL == 'SIM':
        print('REINICIANDO...')
        os.system('python RESGATE.py')
    elif REINICAR_FINAL == 'SS':
        print('REINICIANDO...')
        os.system('python RESGATE.py')
    else:
        print('FECHANDO...\n')
        time.sleep(2)
        os.system('TASKKILL /IM python.exe /T /F')
    
###########################################################################################################
################################################# LINUX ###################################################
###########################################################################################################    
def linux_processar():
    print(colored("---> EM CONSTRUÇÃO <---", 'red', attrs=['bold']))
    print('')   
    print('                    ############          ############')      
    print('                ##################      ######    ####')      
    print('              ######    ######        ####      ####')        
    print('            ######    ######        ######    ####      ####')
    print('          ######    ######          ####    ####      ######')
    print('        ######      ####            ####      ####  ####  ##')
    print('      ######        ####            ####        ######    ##')
    print('     #####          ######          ####          ##    ####')
    print('  ########      ##    ######        ####                ####')
    print('  ######    ########    ######    ####              ######')  
    print('  ####    #####  #####    ##########      ##############')    
    print('    ####  ####    ######    ######      ##############')      
    print('      ######        ######      ####  ####')                  
    print('        ##            ######      ######')                    
    print('                        ####        ####')                    
    print('                      ########        ####')                  
    print('                    ####  ######        ######')              
    print('        ##############      ######        ######')            
    print('      ##############      ##########        ######')          
    print('    ######              ####    ######        ######')        
    print('  ####                ####        ######        ######')      
    print('  ####    ##          ####          ######        ######')    
    print('  ##    ######        ####            ######        ######')  
    print('  ##  ####  ####      ####              ######        ######')
    print('  ######      ####    ####                ######        ####')
    print('  ####      ####    ######                  #####      ####')
    print('          ####      ####                      ####    ######')
    print('        ####    ######                          ##########')  
    print('        ############                              ######')   
    
###########################################################################################################
############################################ INICIO DO PROGRAMA ###########################################
###########################################################################################################
def main():
    global g_ips, g_ports, connections, lastValidHost, validPort
    carregar_configs()

    print("\nLISTA DE IPS CADASTRADOS:")
    for _, y in g_ips.items():
        print(y)
    print()
    
    ########## GRAVA O IP NAS PORTAS GLOBAIS: lastValidHost e ValidPort ##########
    ########## VERIFICA SE IP E PORTAS SÃO VALIDOS ##########
    ########## RETORNA "" E -1 CASO NÃO ENCONTRE O VALIDADOR ##########
    get_host_ip_address(ports=list(g_ports.values()))
    if(validPort == -1):
        conexao_falhou()
        return
    
    ########## VERIFICA SE O VALIDADOR ESTA NA MESMA REDE INTERNA QUE O COMPUTADOR ##########
    is_same_network = windows_is_local_networked_host(lastValidHost)
    if(not is_same_network):
        conexao_falhou()
        return

    detect_os()

if(__name__ == "__main__"):
    main()
    
###########################################################################################################
"""
    IMPORTANTE:

    ORDEM DE EXECUÇÃO DO SCRIPT:

        -> Carregar globais e constantes do topo do script
        -> detect_os()
        -> get_host_ip_address()
        -> windows_is_local_networked_host()
        -> main()
            -> executar_windows()
            -> get_ssh_session()
               OU
            -> executar_linux()
            -> get_ssh_session()

    Anotações:

        - get_host_ip_address():
            Esta função avalia todos os IPs que estão no arquivo de configuração json. O IP e porta do host
            que responder a solicitação deste terminal ficarão salvos em duas globais: "lastValidHost" e "validPort".
        
        - windows_is_local_networked_host() -> bool:
            Envia um comando ping e, através dele, verifica se o host remoto está na mesma rede interna (IPv4 TTL=64).
            Se estiver na mesma rede interna, retornar verdadeiro.

        - get_ssh_session() -> SSHClient:
            Tenta estabelecer conexão com o host remoto. Se houver sucesso, retorna um objeto de SSHClient que faz
            referência a conexão estabelecida.
        
        - processar_dados_validador(SSHClient):
            Obtém dados do validador.
    
    Observações:

        - Não esquecer de criar uma função equivalente de "windows_is_local_networked_host" para o Linux.
"""
###########################################################################################################
