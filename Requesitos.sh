#########################################################
#                       Inicial                         #
#                                                       #
# Instalação das dependencias e prograsmas adicionais   #
# para rodar o RESGATE                                  #
#                                                       #
#                                                       #
#########################################################

echo 'Inciando Instalação'

sudo apt-get update 
sudo apt-get upgrade
sudo apt-get install cifs-utils
sudo pip install paramiko
sudo pip install openpyxl
sudo pip install termcolor
sudo chmod 777 RESGATE.py
sudo mkdir /media/Servidor
sudo mkdir /media/PRODATA

echo 'Instalação Finalizada'