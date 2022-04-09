##################################################
#                                                #
#       Auditoria de Bakups EXA Soluções         #
#						 #
#  V..1  2019-05-21                             #
#  DeV: Luiz Buffa  e  Augusto Peniel		 #
#  Sup: Paulo Renato e Raimundo Sergio           #
#  Emp: EXA Soluções                             #
#						 #
##################################################

#Imports do sistema 
import os
import os.path
from oauth2client.client import crypt
## Autenticedores e api google ##
import json  
import gspread
import pygsheets
#from pydrive.auth import GoogleAuth
#from pydrive.drive import GoogleDrive
from oauth2client.client import SignedJwtAssertionCredentials
#from google_drive_downloader import GoogleDriveDownloader as gdd

#################################
from datetime import time, timedelta, date, datetime
import time


##Variaveis globais
#
data_atual1 = date.today()
tabelalinha = "" 
s_con_log = ""
folder_N = 0 


## Função de log para auditoria
#
def report_log(text1):
        global data_atual1
        try:
                arquivo_log = open(os.getcwd() + "/Log_BKP.txt", 'r')
                text = arquivo_log.readlines()
                text.append("\n"+ str(data_atual1)+" >>># "+ str(text1))
                arquivo_log = open(os.getcwd() + "/Log_BKP.txt", 'w')
                arquivo_log.writelines(text)
                arquivo_log.close()  
        except:
                arquivo_log = open(os.getcwd() + "/Log_BKP.txt", 'w+')
                arquivo_log.close() 
         

##Busca, verificação e criação do arquivo de configuração
#verificação da existencia do arquivo de configuração (caso não exista ele será criado)
#divisão em vetor ferramenta e locais de arquivo log 
#
def arq_de_configuracao():
        global tabelalinha
        global s_con_log 
        try:
                arquivo_config = open(os.getcwd() + "/ini.txt", 'r')
                for linha in arquivo_config:
                        
                        linha = linha.rstrip('\n')

                        if '>Linha_Tabela:#' in linha :
                                linha_dados_acesso_spl =  linha.split("#") 
                                tabelalinha = str(linha_dados_acesso_spl[1])
               
                arquivo_config.close()
                
        except FileNotFoundError:
                arquivo = open(os.getcwd() + "/ini.txt", 'w+')   
                texto = arquivo.readlines()
                texto.append(">Linha_Tabela:#\n")       
                arquivo.writelines(texto)

                arquivo.close()
        print(tabelalinha)
        pass

##Busca e processamento de data do sistema 
#busca de data atual do sistema operaçoinal
#processa resultando em data do dia anterior
#
def buscar_data_sys():

   data_formatada = []
   data_atual = date.today()
   data_ontem = data_atual - timedelta(1)
   data_formatada.append(data_ontem.strftime("%d.%m.%Y"))
   data_formatada.append(data_ontem.strftime("%d/%m/%Y"))
   data_formatada.append(data_ontem.strftime("%d/%m/%y"))
   
 
   return data_formatada 

## Função para verificar o tipo de repositório e tratativa conforme
# Divisão entre Nas ou pasta local
#
def Verificação_caminho_repositório(coluna):
        global s_con_log
        status=""
        report_log("Iniciando Verificação de caminho ")
        try: 
                cam_arq = leitura_tabela(coluna,"")
                if cam_arq !="0":
                        os.system(r"NET USE B: /DELETE")
                        if "$" in cam_arq:
                                cam_arq = cam_arq.split("$")                
                                password = leitura_tabela("C","")
                                if password != "0":
                                        password = " /user:ubackup " + password
                                report_log("Montando caminho Externo: "+ cam_arq[0])
                                os.system("NET USE B: " + cam_arq[0]+"$" + password)
                                fille = ("B:"+cam_arq[1]) 
                                report_log("NET USE B: " + cam_arq[0]+"$" +"password")
                                report_log("Caminho apartir do mapeamento: "+ cam_arq[1])
                                report_log("Verificando caminho Externo: "+ fille)
                                status = verif_arq(fille, coluna)

                        elif "#" in cam_arq:
                                cam_arq = cam_arq.split("#")                
                                password = leitura_tabela("C","")
                                if password != "0":
                                        password = " /user:ubackup " + password
                                report_log("Montando caminho Externo: "+ cam_arq[0])
                                os.system("NET USE B: " + cam_arq[0] + password)
                                fille = ("B:"+cam_arq[1]) 
                                report_log(r"NET USE B: " + cam_arq[0] + " /user:ubackup  password")
                                report_log("Caminho apartir do mapeamento: "+ cam_arq[1])
                                report_log("Verificando caminho Externo: " + fille)
                                status = verif_arq(fille, coluna)

                        elif ":\\" in cam_arq :
                                fille =  cam_arq 
                                status = verif_arq(fille,coluna)
                                report_log("Verificando caminho Interno "+ fille)

                        else:
                                report_log("Falha na leitura do caminho do repositorio de arquivo")                        
                                s_con_log = s_con_log + "#Falha na leitura do caminho do repositorio de arquivo#"                        

                        os.system(r"NET USE B: /DELETE") 
                else:
                        report_log(" Falha na Coluna correspondente ao caminho do arquivo, falha na conexão com o reositório ") 
                        s_con_log = s_con_log + "# Falha na Coluna correspondente ao caminho do arquivo, falha na conexão com o reositório#"

        except:
                report_log("Falha na leitura do caminho do repositorio de arquivo") 
                s_con_log = s_con_log + "#Falha na leitura do caminho do repositorio de arquivo#"
 

        return status

##função de verificação dos arquivos por ferramenta
# Filtro para ferramenta conforme coluna usada na taela 
# ferramenta Veeam e SqlBackup
#
def verif_arq(fille, coluna):
        from datetime import time, timedelta, date, datetime
        import time
        global s_con_log
        
        data_atual11 = date.today() - timedelta(1)
        data_atual11 = data_atual11.strftime("%Y-%m-%d")
        
        data_atual12 = date.today() 
        data_atual12 = data_atual12.strftime("%Y-%m-%d")
        status=""

        # Veeam verificação de ARQ 
        if coluna != "I":

                try:
                        x =os.listdir(fille)        
                        nom_arq_vib_fill=[]
                        nom_arq_vbk_fill=[]
                        result=[] 

                        ## Busca do nome dos arquivos .Vbk e .Vib
                        #Os nome dos arquivos são armazenados em vetores para que se possa verificar mais de um arquivo com mesma datação 
                        #
                        # .Vbk
                        cont=0
                        
                        while  cont < int(len(x)) :
                                if ".vbk" in x[cont]:
                                        nom_arq_vbk_fill.append(x[cont])
                                        
                                cont = cont + 1

                        # .Vib
                        cont=0
                        found=0
                        while  cont  < int(len(x)) and  found == 0:
                        
                                if ".vib" in x[cont]:
                                        if data_atual12 in x[cont]: 
                                                nom_arq_vib_fill.append(x[cont])
                                                found = 1
                                        if data_atual11 in x[cont]:
                                                nom_arq_vib_fill.append(x[cont])                                     
                                cont = cont + 1

                                
                        print(nom_arq_vib_fill)
                        report_log(nom_arq_vib_fill)
                        print(nom_arq_vbk_fill)
                        report_log(nom_arq_vbk_fill)


                        ##Verificação do arq .Vib(incremental)
                        cont = 0 
                        while  cont < int(len(nom_arq_vib_fill)):

                                # Verificando Data de criação do arquivo .Vib
                                report_log(fille+"\\"+nom_arq_vib_fill[cont])
                                wd= time.localtime(os.path.getctime(str(fille+"\\"+nom_arq_vib_fill[cont])))

                                ########## tratamento de datas
                                dia=""
                                mes=""
                                if wd.tm_mon < 10:
                                        mes = "0"+str(wd.tm_mon)
                                else:
                                        mes = str(wd.tm_mon)
                                if wd.tm_mday<10:
                                        dia = "0"+str(wd.tm_mday)
                                else:
                                        dia = str(wd.tm_mday)
                                ########### tratamento de datas
                                
                                if (str(wd.tm_year)+"-"+ mes +"-"+ dia) in (fille+"\\"+nom_arq_vib_fill[cont]):
                                        print( "Success Data de criação .Vib")
                                        result.append("Success")
                                else:
                                        result.append("Failed")
                                        print( "Failed Data de criação .Vib")
                                        print(fille+"\\"+nom_arq_vib_fill[cont])
                                

                                # Verifcação da Data de modificação do arquivo .Vib com relação de 3 dias 
                                print("Etrou na verif data de modificação do Vib com relação da data de 3 dias")   
                                wd= time.localtime(os.path.getmtime(str(fille+"\\"+nom_arq_vib_fill[cont])))
                                cont1 = -1
                                while cont1 < 2:  
                                        data_x = date.today() - timedelta(cont1)
                                        data_x = data_x.strftime("%Y-%m-%d")
                                        dia=""
                                        mes=""
                                        if wd.tm_mon < 10:
                                                mes = "0"+str(wd.tm_mon)
                                        else:
                                                mes = str(wd.tm_mon)
                                        if wd.tm_mday<10:
                                                dia = "0"+str(wd.tm_mday)
                                        else:
                                                dia = str(wd.tm_mday)

                                        report_log(str(wd.tm_year)+"-"+ mes +"-"+dia)
                                       
                                        if (str(wd.tm_year)+"-"+ mes +"-"+dia) == (str(data_x) ):
                                                print("Resultado da verif data de modificação do Vib com relação da data de 3 dias: Success")
                                                result.append("Success")
                                                cont1 = 3
                                        else:         
                                                cont1 = cont1 + 1
                                
                                

                                #Verificação do tamanho dos arquivos .Vib
                                # tamanho convertido de bytes para Megabytes
                                print("Verificação de Arqruivo - tamanho do Vib:")
                                report_log("Verificação de Arqruivo - tamanho do Vib:" + str((os.path.getsize((fille+"\\"+ nom_arq_vib_fill[cont])))/131072))
                                if ( (os.path.getsize(str(fille+"\\"+ nom_arq_vib_fill[cont])))/131072) > 40:
                                        report_log( "Success tamanho do Vib:")
                                        print( "Success tamanho do Vib:")
                                        result.append("Success")
                                else:
                                        report_log( "Failed tamanho do Vib:")
                                        print( "Failed tamanho do Vib:")
                                        result.append("Failed")
                                                             

                                cont = cont + 1
                          

                        ##Verificação do arq .Vbk(full)

                        cont = 0 
                        while  cont < int(len(nom_arq_vbk_fill)):

                                # verificado se a data de modificaçãod o bakup full(.Vbk) é igual a data de criação do incremental(.Vib) 
                                wo= time.localtime(os.path.getmtime(str(fille+"\\"+nom_arq_vbk_fill[cont])))
                                wd= time.localtime(os.path.getctime(str(fille+"\\"+ nom_arq_vib_fill[cont])))

                                ############# tratamento de datas
                                dia=""
                                mes=""
                                if wd.tm_mon < 10:
                                                mes = "0"+str(wd.tm_mon)
                                else:
                                                mes = str(wd.tm_mon)
                                if wd.tm_mday<10:
                                                dia = "0"+str(wd.tm_mday)
                                else:
                                                dia = str(wd.tm_mday) 

                                ############### tratamento de datas
                                dia1=""
                                mes1=""
                                if wo.tm_mon < 10:
                                                mes1 = "0"+str(wo.tm_mon)
                                else:
                                                mes1 = str(wo.tm_mon)
                                if wo.tm_mday<10:
                                                dia1 = "0"+str(wo.tm_mday)
                                else:
                                                dia1 = str(wo.tm_mday)  
                                ###############  tratamento de datas        


                                report_log("Comparação de data criação vib com Moodificação vbk: "+str(wo.tm_year)+"-"+ mes1 +"-"+dia1+" "+str(wd.tm_year)+"-"+ mes +"-"+dia)
                                print("Comparação de data criação vib com Moodificação vbk: "+ (str(wo.tm_year)+"-"+ mes1 +"-"+dia1)+" "+(str(wd.tm_year)+"-"+ mes +"-"+dia))
                                count2= -1
                                result2= 0
                                while count2 < 2:
                                        
                                        data_y = date.today() - timedelta(count2)
                                        data_y = data_y.strftime("%Y-%m-%d")
                                        if (str(wd.tm_year)+"-"+ mes +"-"+dia) == (str(data_y) ):
                                                result2 = 1
                                        if result2 == 1:    
                                                result2=1           
                                        else:
                                                result2=0
                                                    
                                        count2 = count2 + 1     

                                if result2 == 1:
                                        print("Sucesso na verificação de comparação de datas criação vib com Moodificação vbk arq veeam")
                                        report_log("Sucesso na verificação de comparação de datas criação vib com Moodificação vbk arq veeam")
                                        result.append("Success")
                                else:
                                        print("Falha na verificação de comparação de datas criação vib com Moodificação vbk arq veeam")
                                        report_log("Falha na verificação de comparação de datas criação vib com Moodificação vbk arq veeam")
                                        result.append("Failed")
                                
                                
                                # Verificação do tamanho do arquivo Full
                                # tamanho convertido de bytes para Megabytes
                                print("Verificação de Arqruivo  - tamanho do Vbk:")
                                report_log("Verificação de Arqruivo  - tamanho do vbk:" + str((os.path.getsize((fille+"\\"+ nom_arq_vbk_fill[cont])))/131072))
                                if ( (os.path.getsize(str(fille+"\\"+ nom_arq_vbk_fill[cont])))/131072) > 40:
                                        report_log( "Success tamanho do Vbk:")
                                        print( "Success tamanho do Vbk:")
                                        result.append("Success")
                                else:
                                        report_log( "Failed tamanho do Vbk:")
                                        print( "Failed tamanho do Vbk:")
                                        result.append("Failed")


                                cont = cont + 1
                                
                        report_log(result) 
                        print(result) 


                        ## Validação do resultado de status
                        print ("Entrando na Validação do resultado de status veeam")
                        cont = 0 
                        while  cont < int(len(result)):
                                if "Failed" == result[cont]:
                                         status="Failed"
                                        
                                else:
                                         status="Success"
                                cont = cont + 1
                        print("Resultado vrificação Arquivo veems: " +status)
                        report_log("Resultado vrificação Arquivo veems: " +status)


                except:
                       report_log("Erro na verificação de integridade de ARQ Veeam" )
                       s_con_log = s_con_log + "#Erro na verificação de integridade de ARQ Veeam" 
                         


        ##Sql_Backup verificação do repositório              
        elif coluna == "I":

                try:
                        global folder_N 
                        x = os.listdir(fille)
                        nom_arq_sql=[]

                        print("Entrando na Verificação de arquivos do SQL_Backup ")
                        report_log("Entrando na Verificação de arquivos do SQL_Backup ")

                        cont=0
                        while  cont < int(len(x)) :

                                wd= time.localtime(os.path.getctime(str(fille +"\\"+ x[cont])))
                                
                                #################### tratamento de dartas
                                dia=""
                                mes=""
                                if wd.tm_mon < 10:
                                                mes = "0"+str(wd.tm_mon)
                                else:
                                                mes = str(wd.tm_mon)
                                if wd.tm_mday < 10:
                                                dia = "0"+str(wd.tm_mday)
                                else:
                                                dia = str(wd.tm_mday)
                                                
                                #################### tratamento de dartas

                                print("Data do arquivo encontrado: "+str(wd.tm_year)+"-"+ mes +"-"+dia)
                                report_log("Data do arquivo encontrado: "+str(wd.tm_year)+"-"+ mes +"-"+dia)
                        
                                if (str(wd.tm_year)+"-"+ mes +"-"+dia) in (str(data_atual11)) :
                                        nom_arq_sql.append(x[cont])     

                                cont = cont + 1
                        

                        if folder_N > len(nom_arq_sql) :
                                status ="Error"
                                report_log("Verificação de Arqruivo Sql_Backup: Error" )
                                print("Verificação de Arqruivo Sql_Backup: Error" )
                        else:
                                status= "Success"
                                report_log("Verificação de Arqruivo Sql_Backup: Success" )
                                print("Verificação de Arqruivo Sql_Backup: Success" )

                except:
                        report_log("Erro na verificação de integridade de ARQ SQL_Backup" )
                        s_con_log = s_con_log + "#Erro na verificação de integridade de ARQ SQL_Backup " 

        else:
                
                report_log("Erro na execução de conferência do repositório" )
                s_con_log = s_con_log + "#Erro na execução de conferência do repositório"        
                
                
        return status

## Função para leitura avuso na tabela
# 
def leitura_tabela(coluna,dataz):
             global data_atual1
             global tabelalinha
             global s_con_log
             va_leitura = ""
             if dataz == "":
                        dataz = data_atual1 

             try:
                cout = 0   
                while cout < 4 : 
                        
                        cout = cout + 1
                        json_key = json.load(open(os.getcwd() + "/teste2-55a2a4ccc47b.json", "r")) # Arquivo de credenciais json
                        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
                        credentials = SignedJwtAssertionCredentials(json_key['client_email'], json_key['private_key'].encode(), scope) 
                        
                        
                        
                        file = gspread.authorize(credentials)                  
                        sheet = file.open("teste1").worksheet(str(dataz))
                        
                        

                        va_leitura = sheet.acell(str(coluna) + tabelalinha ).value
                        
                        if va_leitura =="0":
                                report_log(" leitura de dados na tabela: \'NULL\'")
                        else:
                                if coluna == "C":
                                                report_log(" leitura de dados na tabela: Password ")
                                else:
                                                report_log(" leitura de dados na tabela: " + "\'" + va_leitura + "\' ") 
                

             except:
                report_log(" Erro, na leitura de informações da tabela")
                s_con_log = s_con_log + "#Erro, na leitura de informações da tabela"
              
             return va_leitura
                

## Função de gravação na tabela 
#
def gravacao_tabela(coluna,valor):
                global data_atual1
                global tabelalinha

                

                json_key = json.load(open(os.getcwd() + "/teste2-55a2a4ccc47b.json", "r")) # Arquivo de credenciais json
                scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

                credentials = SignedJwtAssertionCredentials(json_key['client_email'], json_key['private_key'].encode(), scope) 

                file = gspread.authorize(credentials) 
                sheet = file.open("teste1").worksheet(str(data_atual1))
                celula = str(coluna) + tabelalinha

                sheet.update_acell(celula, str(valor))

                if valor=="":
                        report_log(" Gravação de dados na tabela "+ "\'NULL\'")
                else:
                        report_log(" Gravação de dados na tabela "+ "\'"+valor+"\'")
          

## Processamento do vetor da função de configuração do ini 
#para cada ferramenta é acionada uma função para a analise do log 
#
def proces_vetor_config():
     global s_con_log

     # tratamento de Data 
     buscar_data = date.today()
     buscar_data = buscar_data - timedelta(1)
     buscar_data = buscar_data.strftime("%Y-%m-%d")
      
     try:
                #Limpesa das colunas de status 
                gravacao_tabela("K","")
                gravacao_tabela("L","") 
                gravacao_tabela("M","") 
                gravacao_tabela("N","")
                
                                
                #Veeam
                report_log("############## Verificação Veeam ##############")
                if leitura_tabela("D","") != "0" :
                        
                        x = processamento_log_veeam(leitura_tabela("D",""))
                        report_log(x)
                        z = Verificação_caminho_repositório("G")
                        report_log(z)
                        
                        
                        if x=="":
                                print(buscar_data)
                                gravacao_tabela("L",leitura_tabela("L",buscar_data)+" (Copy)")
                                report_log("Status Veeam Status copy")     
                                gravacao_tabela("K","#Veeam Status copy")
             
                        else:
                                if z=="":
                                        report_log("Status Veeam: Failed")     
                                        gravacao_tabela("L","Failed")    
                                else:
                                        report_log("Status Veeam: " + str(x))     
                                        gravacao_tabela("L",str(x))

                report_log("############## Fim Verificação Veeam ##############")                

                
                #Veeam B&R
                report_log("############## Verificação Veeam B&R ##############" )
                if leitura_tabela("E","") != "0" :
                        
                        x = processamento_log_veeam(leitura_tabela("E",""))
                        report_log(x)
                        z = Verificação_caminho_repositório("H")
                        report_log(z)

                        if  x == "":
                                
                                print(buscar_data)
                                gravacao_tabela("M",leitura_tabela("M",buscar_data)+" (Copy)")
                                report_log("Status Veeam B&R Status copy")     
                                gravacao_tabela("K","#Veeam B&R Status copy")       
                        else:
                                if z=="":
                                        report_log("Status Veeam: Failed")     
                                        gravacao_tabela("M", "Failed")
                                else:
                                        report_log("Status Veeam B&R: " + str(x))     
                                        gravacao_tabela("M",str(x))

                report_log("############## Fim Verificação Veeam B&R ##############" )    

                                
                #SQL_Backup
                report_log("############## Verificação SQL_Backup ############## " )
                if leitura_tabela("F","") != "0" :

                        x = processamento_log_sql_backup(leitura_tabela("F",""))
                        report_log(x)
                        z = Verificação_caminho_repositório("I")
                        report_log(z)

                        if  x == "0":
                                gravacao_tabela("N",leitura_tabela("N",buscar_data)+" (Copy)")
                                report_log("Status SQL_BACKUP Status copy")     
                                gravacao_tabela("K","#SQL_BACKUP Status copy")       
                        else:
                                if z=="" :
                                        report_log("Status Sql_Backup \'ERROR\', falha na conexão com o reositório")     
                                        gravacao_tabela("N", "Error")
                                else:
                                        report_log("Status Sql_Backup: " + str(x))     
                                        gravacao_tabela("N",str(x))


                report_log("############## Fim Verificação SQL_Backup ##############" )             
     except:
        report_log("Erro de leitura do arquivo log")
        s_con_log = s_con_log +"#Erro de leitura do arquivo log " 


## Função de analise do arquivo log da ferramenta Postgre_SQL 
#
def  processamento_log_sql_backup(local_log):
         global folder_N 
         global s_con_log
         buscar_data = buscar_data_sys()
         info_log = []
         info="0"

         try:
                 arquivo = open(local_log, "r")
                 x = 0

                 for linha in arquivo: 
                        if "run" in linha:
                                if (str(buscar_data[1])  in  linha) or (str(buscar_data[2])  in  linha):
                                     x=1
                        if "SUMMARY:" in linha:
                                 x = x+3
                        if "DETAILED LOG:" in linha :
                                 x = 0
                        if x == 4 :
                                if "Database" and "Google Drive" in linha:

                                        if "Success" in linha:
                                                info_log.append("Success")
                                        elif "Failed" in linha:
                                                info_log.append("Failed")
                                        else:
                                                info_log.append("Error")

                                elif "Database" and "FTP" in linha:
                                        
                                        if "Success"  in linha:
                                                info_log.append("Success")
                                        elif "Failed" in linha:
                                                info_log.append("Failed")
                                        else:
                                                info_log.append("Error")

                                elif "Database" and "Folder" in linha :

                                        if "Success"  in linha :

                                                info_log.append("Success")

                                                #contador de quantidade de arq para função  de verificação
                                                ################                                         
                                                folder_N = folder_N + 1 
                                                ################

                                        elif "Failed" in linha:
                                                info_log.append("Failed")
                                        else:
                                                info_log.append("Error")
                                
                                elif "Database" and "Dropbox" in linha :

                                        if "Success"  in linha :
                                                info_log.append("Success")
                                        elif "Failed" in linha:
                                                info_log.append("Failed")
                                        else:
                                                info_log.append("Error")
               
                 arquivo.close()

                 for item in info_log :
                         if "Success" in item:
                                 info = str(item)
                         elif "Failed" in item:
                                 info = str(item)
                         elif "Error" in item:
                                 info = str(item)
                         

                 
                 
         except FileNotFoundError:  

                 report_log("Erro,Falha na leitura do log  SQlBakupAndFtp")
                 s_con_log = s_con_log + "#Erro,Falha na leitura do log  SQlBakupAndFtp"
                   


         return info


##Processamento e pesquisa de informação dentro do arquivo log 
#
def  processamento_log_veeam(local_log):
        global s_con_log
        buscar_data = buscar_data_sys()
        linha_log=""
        idot="*"
        linha_log_dados = []
        report_log(local_log)
        try:      
                arquivo = open(local_log, 'r') 
                for linha in arquivo:
                        if str(buscar_data[0]) in linha :
                            if "[Session] Id" in linha:
                                 linha_log_dados = linha.split("'")
                                 idot = linha_log_dados[1] 
                                 
                        if 'Job session' in linha:    
                            if idot in linha:
                                linha_log = linha
                                if 'Success' in linha_log:
                                        linha_log = "Success"
                                elif 'Warning' in linha:
                                        linha_log = "Warning"
                                elif 'Failed' in linha:
                                        linha_log = "Failed"
                               
                arquivo.close()
                 
               
        except FileNotFoundError:
                
                report_log("Erro,Falha na leitura do log Veeam")
                s_con_log = s_con_log + "#Erro,Falha na leitura do log Veeam"
    
        
        return linha_log

##Processamento do arq de configuração.
#
def  hora_exec():
   gravacao_tabela("O","")
   hora_atual = datetime.now()
   wd = datetime.weekday(hora_atual)
   days= ["Segunda","Terça","Quarta","Quinta","Sexta","Sabado","Domingo"]
   gravacao_tabela("O", hora_atual.strftime("%H:%M:%S ") + days[wd])
   
   pass 

##Setagem de flag de execução
#
def set_falg_exec():
        flagexec = 0
        
        flagexec = int(leitura_tabela("P",""),10)
        
        if flagexec < 3 :
                flagexec = flagexec + 1 
                gravacao_tabela("P",str(flagexec))                   
        pass
        
        
## Verificação geral de execução do software
#
def ver_soft_exec():
        global data_atual1
        global tabelalinha
        global s_con_log
        va_leitura = []
        va_leitura1 = [] 
        retorno=None
        dataz = data_atual1
        count = 0 
        json_key = json.load(open(os.getcwd() + "/teste2-55a2a4ccc47b.json", "r")) # Arquivo de credenciais json
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        credentials = SignedJwtAssertionCredentials(json_key['client_email'], json_key['private_key'].encode(), scope)
         
                
                
        file = gspread.authorize(credentials)                  
        sheet = file.open("teste1").worksheet(str(dataz)) 
        try:
                all_cells = sheet.range("L" + tabelalinha +":"+"O"+tabelalinha)
                for cell in all_cells:
                        if cell.value != "":
                                va_leitura.append("1")
                        else:
                                va_leitura.append("0")        
                
                all_cells2 = sheet.range("D" + tabelalinha +":"+"F"+tabelalinha)
                for cell in all_cells2:
                        if cell.value != "0":
                                va_leitura1.append("1")
                        else:
                                va_leitura1.append("0")  
                va_leitura1.append("1")     
                print(va_leitura)
                repete = False 
                while count <= 3 and repete != True: 
                        if va_leitura1[count] == va_leitura[count] : 
                                retorno = 1
                                
                        else:
                                repete = True 
                                retorno = None 
                                
                        count = count + 1
                print(va_leitura1)  
                           
        except:
                        report_log(" Erro, na leitura de informações da tabela, cheagem de campos de status.")
                        s_con_log = s_con_log + "#Erro, na leitura de informações da tabela, checagem de campos de status."
                
        return retorno

##Função main(principal) do software
# 
def main():
           global tabelalinha
           global s_con_log
   
           #Processamento do arq de configuração.
           arq_de_configuracao()
   
           report_log("############################################################################################")
           report_log("############################################################################################")
        
           verificador = ver_soft_exec()
           while  verificador == None :
                   
           #Processamento da verificação.   
                if tabelalinha != "" :
                        if tabelalinha != "1" :
                                proces_vetor_config()
                                gravacao_tabela("K",s_con_log)
                                        
                #Gravação do horario de execução do software na tabela 
                hora_exec()
                report_log("############################################################################################")
                #chamanda da veirficação final de execução do software
                verificador = ver_soft_exec()
                
           #Comando de desmontagem Para evitar que fique qualquer conexão externa criada em aberto
           os.system(r"NET USE B: /DELETE")
   
           #Setagem da ultima flag
           set_falg_exec()              
           
           return 0   
main()
os.system("taskkill /F /IM Verificador.exe")
os.system("taskkill /F /IM net.exe")