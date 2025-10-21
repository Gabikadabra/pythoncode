#New Reportistica ---
###Codice da completare
#Estrae elenco dipendenti 
# #def EstraeRepGMRE(fTest,fListDip,fListRep,fNomeFileOut):
#-----Imposta le librerie (vengono richiamate funzioni presenti in altri moduli)
#Legge la libreria Pandas per il DataFrame/gestione tabelle
from Util_leggeDir import dir_dict
from Util_leggeParams import parm_dict,listfilter_ODAG,listfilter_BDO,listfilter_MAP,listfilter_Task


import pandas as pd
import numpy as np #per sostituire i valori NA	from Funz.Util_wDir import RepInit0	
	#import datetime   #importare libreria per la data del giorno
#Clean non numeric values
import re
#Remove chars not numeric form column
def remove_chars(s):
#2025.04.08 begin
#   return re.sub("[^0-9,]+","",str(s))  # to remove all non alphabets & numbers & commas & dots  re.sub(r"[^a-zA-Z0-9,.]+", '', s)
    a= re.sub("[^0-9,.]+","",str(s))  # to remove all non alphabets & numbers & commas & dots  re.sub(r"[^a-zA-Z0-9,.]+", '', s)
    a=a.strip()
    return a.replace(",",".")
from datetime import timedelta,datetime  
from read_Interni import df_RisInt_byCID, df_RisInt_byName
from read_IFproc import df_IFprc,df_BDOvarprc
from read_RdiBdo_cycle import df_RDI_park, df_RDI, df_RDI_old, df_RDI_port
#from read_ import df_RDI_port
from read_BdoLin import df_BDO

from Util_logging import logger
wPythProc="elab_If_Rdi"
import timeit  #funzione timeit per valutare durata estrazione
tempoInz=timeit.default_timer()


#get information from park mail
#df_RDI_port = pd.concat([df_forn_RTI, df_forn_pos, df_forn_subf,df_forn_add], ignore_index=True, sort=False)

df_RDI_port=pd.merge(df_RDI_port,df_IFprc[["RDI","StatoIF_byPMO","chkIF_byPMO","Note_PMO","BdO_byPMO-BO","Stato_elab","Data_lavorazione_byPMO-BO","Descrizione RDI","ROI","Dip"]],left_on="RDI", right_on="RDI",how="outer",suffixes=('', '_IF'))
df_RDI_port=pd.merge(df_RDI_port,df_RDI_park[["RDI","Macrocl_nbr","ODAGnbr","ROI","ROI_dip","ODAG","Macrocl_txt"]],left_on="RDI", right_on="RDI",how="outer",suffixes=('', '_park'))

#get information from RDI pending
df_RDI_port=pd.merge(df_RDI_port,df_BDO[["RDI","DocAcq","StatoDocAcq","ROI","ROI_dip", "Forn_val","DataCreaz","ODAG", "ODAGnbr","Macrocl_txt","Macrocl_nbr","Descrizione DocAcq"]],left_on="RDI", right_on="RDI",how="left",suffixes=('', '_BDO'))
#exclude RDI non strt with "2"
df_RDI_port=df_RDI_port[df_RDI_port['RDI'].astype(str).str.startswith('2')]
df_RDI_port = df_RDI_port[~df_RDI_port["RDI"].isin(df_RDI_old['Numero RDI'])]

#df_RDI_port["Importo Richiesta"]=df_RDI_port["Importo Richiesta"].apply(remove_chars);df_RDI_port["Importo Richiesta"]=df_RDI_port["Importo Richiesta"].apply(pd.to_numeric,errors="coerce").fillna(0)

#df_RDI_port["DataCreaz"]=pd.to_datetime(df_RDI_port["DataCreaz"], format='%d.%m.%Y')

df_RDI_port["Data_lavorazione_byPMO-BO"]=pd.to_datetime(df_RDI_port["Data_lavorazione_byPMO-BO"], format='%d.%m.%Y')

df_RDI_port['ODAG'] = np.where(df_RDI_port['ODAG'].notnull(), df_RDI_port['ODAG'],   #when... then
#                 np.where(df_RDI_port['ODAG_park'].notnull(), df_RDI_port['ODAG_park'],  #when... then
                  np.where(df_RDI_port['ODAG_BDO'].notnull(), df_RDI_port['ODAG_BDO'],  #when... then
                    ""))   
df_RDI_port['ODAGnbr'] = np.where(df_RDI_port['ODAGnbr'].notnull(), df_RDI_port['ODAGnbr'],   #when... then
#                 np.where(df_RDI_port['ODAGnbr_park'].notnull(), df_RDI_port['ODAGnbr_park'],  #when... then
                  np.where(df_RDI_port['ODAGnbr_BDO'].notnull(), df_RDI_port['ODAGnbr_BDO'],  #when... then
                    ""))   
df_RDI_port['Macrocl_txt'] = np.where(df_RDI_port['Macrocl_txt'].notnull(), df_RDI_port['Macrocl_txt'],   #when... then
#                  np.where(df_RDI_port['Contratto Macroclasse_BDO'].notnull(), df_RDI_port['Contratto Macroclasse_BDO'],
                  np.where(df_RDI_port["Macrocl_txt_BDO"].notnull(), df_RDI_port["Macrocl_txt_BDO"],                             #when... then
                    "")) 
df_RDI_port['BDO RDI'] = np.where(df_RDI_port['DocAcq'].notnull(), df_RDI_port['DocAcq'],   #when... then
                 np.where(df_RDI_port['BdO_byPMO-BO'].notnull(), df_RDI_port['BdO_byPMO-BO'], 
                    ""))
df_RDI_port['Descrizione RDI'] = np.where(df_RDI_port['Descrizione RDI'].notnull(), df_RDI_port['Descrizione RDI'],   #when... then
                 np.where(df_RDI_port['Descrizione RDI_IF'].notnull(), df_RDI_port['Descrizione RDI_IF'], 
                    ""))
df_RDI_port['BDO Stato RDI'] = np.where(df_RDI_port['StatoDocAcq'].notnull(), df_RDI_port['StatoDocAcq'],   #when... then
                 np.where(df_RDI_port['Stato_elab'].notnull(), df_RDI_port['Stato_elab'], 
                    ""))
df_RDI_port['Dip RDI'] = np.where(df_RDI_port['ROI_dip'].notnull(), df_RDI_port['ROI_dip'],   #when... then
                 np.where(df_RDI_port['ROI_dip_BDO'].notnull(), df_RDI_port['ROI_dip_BDO'], 
                 np.where(df_RDI_port['ROI_dip_park'].notnull(), df_RDI_port['ROI_dip_park'], 
                                np.where(df_RDI_port['Dip'].notnull(), df_RDI_port['Dip'],
                    ""))))
df_RDI_port['ROI'] = np.where(df_RDI_port['ROI'].notnull(), df_RDI_port['ROI'],   #when... then
                 np.where(df_RDI_port['ROI_BDO'].notnull(), df_RDI_port['ROI_BDO'], 
                 np.where(df_RDI_port['ROI_park'].notnull(), df_RDI_port['ROI_park'], 
#                 np.where(df_RDI_port['Responsabile Operativo'].notnull(), df_RDI_port['Responsabile Operativo'], 
               np.where(df_RDI_port['ROI_IF'].notnull(), df_RDI_port['ROI_IF'], 
                    ""))))
#rename column labels
#df_RDI_port=df_RDI_port.rename(columns={"Contratto Gara":"Contratto Gara prec", "Contratto Gara RDI":"Contratto Gara","Numero Contratto Gara":"Numero Contratto Gara prec","Numero Contratto Gara RDI":"Numero Contratto Gara","Contratto Macroclasse":"Contratto Macroclasse prec","Contratto Macroclasse RDI":"Contratto Macroclasse","ROI":"ROI rich","ROI RDI":"ROI" })
#df_RDI_port=df_RDI_port.rename(columns={"ROI":"ROI rich","ROI RDI":"ROI" })
#select recent RDIs and without BDO rilasciato
wFilterRDI=[["Completato","BDO Rilasciato"]]

df_RDI_port=df_RDI_port.loc[(df_RDI_port["RDI"].values>2024)&(df_RDI_port["BDO RDI"]=="")]

tempoFin = timeit.default_timer()
logger.info('end reading ' + str(df_RDI_port.shape[0])+" Durata processo: " + wPythProc +" - dim: "+str(df_RDI_port.memory_usage().sum())+" Durata processo: "  + str(round(tempoFin-tempoInz,3))+ " Dataframe dimensione " + str(df_RDI_port.memory_usage(deep=True)) )


pass