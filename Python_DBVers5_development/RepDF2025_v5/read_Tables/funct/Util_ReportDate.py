#Data di estrazione dei report
#Import librerie
import os
import pandas as pd
from datetime import datetime
import datetime 
#modulo per determinare se file è vuoto
from .Util_FileEmpty import file_is_empty
wDtFormat = "%d/%m/%Y"  
def GetRptDate(wdirInp):
    #File csv utilizzato in input
    wfileCSV="Elab_DataRepPDC"
    #se il file è vuoto non esegue l'elaborazione
    if file_is_empty(wdirInp+"\\"+wfileCSV+".csv"): 
            return(r"01/01/2024")
    df_rep = pd.read_csv(wdirInp+"\\"+wfileCSV+".csv", sep=';',on_bad_lines='skip',dtype={"Data Estrazione":str} )
    df_rep["Data Estrazione"]=pd.to_datetime(df_rep["Data Estrazione"]); df_rep["Data Estrazione"]=df_rep["Data Estrazione"].dt.strftime(wDtFormat)
    wDate=df_rep['Data Estrazione'].iloc[0]
    print(df_rep.head)
    print(wDate)
    #Trasforma la data in gormato gg-mm-aaaa
    
    #wDaten=wDate[9:10]+"/"+wDate[6:7]+"/"+wDate[0:4]
    print(wDate)
   # wDate_n=datetime.strptime(wDate, '%Y-%m-%d %H:%M:%S').strftime('%S:%M:%H %d-%m-%Y')
    #df_rep["Data Estrazione"]=df_rep["Data Estrazione"].str.split(" ").str[0]
    return(wDate)
#Estrae il progressivo delle colonne per trascrizione in Excel 
def GetAlfb(wListCol):
#Elenco colonne
    wListAlfb=["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z","AA","AB","AC","AD","AE","AF","AG","AH","AI","AJ","AK","AL","AM","AN","AO","AP","AQ","AR","AS","AT","AU","AV","AW","AX","AY","AZ","BA","BB","BC","BD","BE","BF","BG","BH","BI","BJ","BK","BL","BM","BN","BO","BP","BQ","BR","BS","BT","BU","BV","BW","BX","BY","BZ","CA","CB","CC","CD","CE","CF","CG","CH","CI","CJ","CK","CL","CM","CN","CO","CP"]
#inizializza la Lista
    wListOut=[]
#Scorre la lista e crea nuova lista con le posizioni su Excel
    for wItem in wListCol: wcoln=wListAlfb.index(wItem);wListOut.append(wcoln)
    return(wListOut)