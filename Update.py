#Imports do sistema 
import os
import os.path
## Autenticedores e api google ##
import json 
import gspread
import pygsheets
from oauth2client.client import SignedJwtAssertionCredentials
#################################
from datetime import time, timedelta, date, datetime
import time
from google_drive_downloader import GoogleDriveDownloader as gdd

###############
### Globais ###
###############
tabelalinha = 0
data_atual1 = date.today()
###############


## conexão com a tablela 
# gravação de informações e leitura( função genérica)
def conexao_tabela(coluna,linha,tabela,Valor,leitgrav):
    global data_atual1
    
    # Arquivo de credenciais json
    json_key = json.load(open('c:/Backup-EXA/Data/Verificador/teste2-55a2a4ccc47b.json', 'r')) 
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = SignedJwtAssertionCredentials(json_key['client_email'], json_key['private_key'].encode(), scope) 
    file = gspread.authorize(credentials) 
    sheet = file.open("teste1").worksheet(tabela)
    
    
    #Condicional de leitura 
    if leitgrav == "L":
        valor = sheet.acell(str(coluna)+linha).value
        return valor
    
    #condicional de gravação 
    if leitgrav == "G":
        sheet.update_acell(coluna+linha, str(Valor))
        pass
    
## Função que realiza o Update do software
#
def Update():
    
    gdd.download_file_from_google_drive(file_id = conexao_tabela("E","2","Manutencao","","L"),dest_path = 'C:/Backup-Exa/Data/Verificador/Verificador.exe', overwrite = True)
    
    gdd.download_file_from_google_drive(file_id = conexao_tabela("F","2","Manutencao","","L"),dest_path = 'C:/Backup-Exa/Data/EXAgent/EXAgent.exe', overwrite = True,)  
    
    pass

##Import de informações do ini
# ini se encontra na pasra do "Verificador"
def import_informacoes():
    
    ## Globais importadas ##
    global tabelalinha
    ########################
    
    try:
        
        arquivo_config = open('c:\Backup-EXA\Data\Verificador\ini.txt', 'r')
        for linha in arquivo_config:
            linha = linha.rstrip('\n')
            if '>Linha_Tabela:#' in linha:
                linha_dados_acesso_spl = linha.split("#")
                tabelalinha = str(linha_dados_acesso_spl[1])
        arquivo_config.close()
        
    except FileNotFoundError:
        arquivo=open('c:\Backup-EXA\Data\Verificador\ini.txt', 'w+')
        texto=arquivo.readlines()
        texto.append(">Linha_Tabela:#\n")
        arquivo.writelines(texto)
        arquivo.close()
    

    pass

## verifica flag de update na tabela manutenção 
#
def flag_de_atualizacao():
    global data_atual1
    global tabelalinha
    
    if conexao_tabela("C",tabelalinha,"Manutencao","","L") == "S":
        Update()# Realiza o update
        conexao_tabela("C",tabelalinha,"Manutencao","N","G")    
    pass

## Matar processo do Verificador 
#
def kill_process():
    os.system("taskkill /F /IM Verificador.exe")
    os.system("taskkill /F /IM net.exe")
    os.system("taskkill /F /IM EXAgent.exe")
    
    pass       

def reiniciar():
    os.system("c:/Backup-EXA/Data/EXAgent/EXAgentBkp.bat")
    
    
## Função principal do sistema
#
def main():
    
    kill_process() #Corta processos para não ocorrer conflitos 
    import_informacoes() #Importa informações parara realizar o update 
    flag_de_atualizacao() # Verifica flag se true, ele realiza aexecução do update
    time.sleep(40)
    reiniciar()
   
    
    pass
main()
