#################################
import schedule
import time
import os
import json  
import gspread
import pygsheets
from oauth2client.client import crypt
from oauth2client.client import SignedJwtAssertionCredentials
from datetime import timedelta, date, datetime
#################################

data_atual1 = date.today()
tabelalinha = "" 

## Report log do agente 
#
def report_log(text1):
        global data_atual1
        try:
                arquivo_log = open(os.getcwd() + "/Log_EXAgent.txt", 'r+')
                text = arquivo_log.readlines()
                text.append(str(text1))
                arquivo_log.writelines(str(data_atual1) + ">> " + text)
                arquivo_log.close()  
        except:
                arquivo_log = open(os.getcwd() + "/Log_EXAgent.txt", 'w+')
                arquivo_log.close() 
        

##Import de informações do ini
# ini se encontra na pasra do "Verificador" 
def arq_de_configuracao():
        global tabelalinha
         
        try:
                arquivo_config = open("c:/Backup-EXA/Data/Verificador/ini.txt", 'r')
                for linha in arquivo_config:
                        
                        linha = linha.rstrip('\n')

                        if '>Linha_Tabela:#' in linha :
                                linha_dados_acesso_spl =  linha.split("#") 
                                tabelalinha = str(linha_dados_acesso_spl[1])
               
                arquivo_config.close()
                
        except FileNotFoundError:
                arquivo = open(os.getcwd() + "c:/Backup-EXA/Data/Verificador/ini.txt", 'w+')   
                texto = arquivo.readlines()
                texto.append(">Linha_Tabela:#\n")       
                arquivo.writelines(texto)
                arquivo.close()
                

## conexão com a tablela 
# gravação de informações e leitura(função genérica)
def conexao_tabela(coluna,linha,tabela,Valor,leitgrav):
    global data_atual1
    
    # Arquivo de credenciais json
    json_key = json.load(open('c:/Backup-EXA/Data/EXAgent/teste2-55a2a4ccc47b.json', 'r')) 
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = SignedJwtAssertionCredentials(json_key['client_email'], json_key['private_key'].encode(), scope) 
    file = gspread.authorize(credentials) 
    sheet = file.open("teste1").worksheet(str(tabela))
    
    
    #Condicional de leitura 
    if leitgrav == "L":
        valor  = sheet.acell(coluna+linha).value
        report_log("Realizado conexão Leitura:" + valor)
        return valor
    
    #condicional de gravação 
    if leitgrav == "G":
        sheet.update_acell(coluna+linha, str(valor))
        report_log("Realizado conexão Gravação:" + str(valor))
        return 0



## Execução do job aartir do loop 
#
def job():
  global data_atual1
  global tabelalinha
  
  tempo = datetime.today()
  tempo = tempo.strftime("%H")
              
              
  #Condicional de execução do verificador.exe 
  if int(conexao_tabela("J",tabelalinha,data_atual1,"","L"),10) == int(tempo,10):
        if int(conexao_tabela("P",tabelalinha,data_atual1,"","L"),10) < 3:  
                print("Realizado inicialização do Verificador")
                time.sleep(2)
                os.system("start c:/Backup-EXA/Data/Verificador/VerificadorBkp.bat")
                
                

def job2():
  global tabelalinha

  print("Realizado inicialização do Update")
            
  if conexao_tabela("C",tabelalinha,"Manutencao","","L") == "S":
                os.system("taskkill /F /IM Verificador.exe")
                os.system("taskkill /F /IM net.exe")
                report_log("Realizado inicialização do Update")
                os.system("c:/Backup-EXA/Data/Update/Update.exe")
    
def main():
  arq_de_configuracao()

  ## Loop 
  schedule.every(7).minutes.do(job2)
  schedule.every(10).minutes.do(job)
  
  while True:
    schedule.run_pending()
    time.sleep(1)
  ####### 
  
main()
