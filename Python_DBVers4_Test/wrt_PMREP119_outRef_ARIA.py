#Genera file excel con mail e tabella html filtrata per dipartimento 
"""Mail Invio Verbali ROI DEC e RUP"""
#Legge la libreria Pandas per il DataFrame/gestione tabelle
import pandas as pd
import numpy as np #per sostituire i valori NA	from Funz.Util_wDir import RepInit0	
from datetime import timedelta,datetime  

import html2text
import xlsxwriter
import os
import sys
import re



from Util_leggeDir import dir_dict,wEnv

wrk_RptDir=dir_dict["rptDir"]

wrk_home_dir=dir_dict["homeDir"]
##	Gestione DIRECTORY CORRENTE (contestualizzazione cartella python sulla base dell’utente)
"""NOME FILE in INPUT """
#wNamFile = "Verbali in firma ROI.xlsx
f_reptname="PMREP119 VerbApprov" #"PMRep117-test"

wrk_dir_inp=wrk_RptDir

#wRept_output=wrk_RptDir+"\\"+f_reptname+".xlsx"   
# Posiziona nella cartella di lavoro
#wrk_dir_inp = wrk_home_dir+r"\PMO Shared\Stage\BRampoldi"  #Stage\BRampoldi"
wrk_dir_out = wrk_home_dir+r"\Test PMO-report - Documenti\General\MailDestinatari"


"""NOME COLONNE utilizzate per selezione ed invio """
wcol_name_Dip = "RisInt_Dip" #Nome colonna utilizzata per le strutture
wcol_name_mail ="E-mail" #Nome colonna utilizzata per le mail
wcol_name_PMOmail ="E-mail_pmo" #Nome colonna utilizzata per le mail PMO

#-----Predisposizione delle variabili che verranno utilizzate nel corso dell'elaborazione
wDt_now = pd.to_datetime('now') # Restituisce la data di oggi
#Definire nome file output
nome_xls= "Verbali"		
fNomeFileOut="PMREPOutp_test"

"""TESTO MAIL standard da utilizzate per selezione ed invio """
#TESTO MAIL - prima parte
wmail_txt_inz="<p>Buongiorno,</p><br><p>sulla base dell'estrazione di ieri sera risultano alla vostra approvazione i seguenti verbali (di apertura/chiusura/SAL) relativi ai BDO riportati nella tabella.</p><br><p>I verbali con l'approvazione mancante del referente riportato nella colonna [In approvaz.(Aria)].</p>"
#TESTO MAIL - seconda parte
wmail_txt_fin="<p>Qualora abbiate già provveduto potete ignorare la presente comunicazione.</p><br><br><p>Cordiali saluti</p><br><strong>____________________________________________________<br></strong><br/></p>"

FirmaFC = "<p>Fabio Caneri<br/><br/>Program Management Office</p><br/><p>Cell:+39 3358478739</p><br/><p>[#PwrAut#]</p><br/>"
FirmaEF ="<p>Elisabetta Filippone<br/>Program Management Office</p><p>External</p><br/><p>Referente ARIA S.p.A.: Fabio Caneri, fabio.caneri@ariaspa.it</p><br/><br/><p>Cell:+39 3921094646</p><br/><p>Mail: elisabetta.filippone@ext.ariaspa.it</p><br/><p>[#PwrAut#]</p><br/>"
FirmaBR ="<p>Betsabea Rampoldi<br/><br/>Program Management Office</p><br/><p>Mail: betsabea.rampoldi@ariaspa.it</p><br/><p>[#PwrAut#]</p><br/>"

"""POSISIZONE COLONNE da inserire nella tabella html"""
wColfromExcel=[0,2,11,4,6,7,5,8,9,10]

"""OGGETTO della mail"""
wmail_oggetto="Verbali in approvaz."

if wEnv == "FCan":
	wFirma=FirmaFC
elif wEnv=="BRam":
	wFirma=FirmaBR
elif wEnv=="EFil":
	wFirma=FirmaEF
else:
	wFirma=""

wmail_end="<p>Ai sensi del Regolamento (UE) 2016/679, Le informazioni contenute in questo messaggio e ogni documento o file ad esso allegato sono riservati e confidenziali. Il loro utilizzo è consentito esclusivamente al destinatario del messaggio o a diversa persona da questo autorizzata, per le finalità indicate nel messaggio medesimo. Qualora Lei non fosse la persona cui il presente messaggio è destinato, La invitiamo a eliminarlo dal Suo Sistema e a distruggere le varie copie o stampe, dandocene gentilmente comunicazione. Ogni utilizzo improprio è contrario ai principi del Regolamento (UE) 2016/679. Aria S.p.A. opera in conformità al Regolamento (UE) 2016/679 citato.</p><br/><p>The information contained in this message (including attachments), for the Regulation (EU) 2016/679is private and confidential and is only intended for the person to whom it is addressed. If the reader of this message is not the intended recipient or the employee or agent responsible for delivering the message to the intended recipient, or you have received this communication in error, please be aware that any dissemination, distribution or duplication is strictly prohibited, and may be illegal. Please notify us immediately and delete all copies from your mailbox and other archives. Thank you for your cooperation.</p>" 
#crea parte finale della mail 
wmail_txt_fin=wmail_txt_fin+wFirma+wmail_end


wfirstloop=True

#Nome del file utilizzato in output
wNamTable= "Tabella1" #Nome della tabella che verrà utilizzata nel file Excel



#creare dataframe (inizialmente vuoto) per inserire mail
col_names=["Dip","TestoMail","Oggetto","To_Dest","Cc_Dest"]
df_tabmail=pd.DataFrame(columns=col_names)
df_work=reptApprovVerb=pd.read_excel(wrk_RptDir+"\\"+f_reptname+".xlsx",sheet_name="VerbSAL_ApprovARIA", dtype={"BDO": str,"Periodo": str,"ODAGforn":str,"Data Creazione Bdo":str,	"Data Approvazione":str,	"Data decorrenza":str,	"Data scadenza":str,"Tipologia Fornitura":str,"Stato Testata Ordine":str}, decimal=',')

#df_work=pd.read_excel(wrk_dir_inp+"\\"+wNamFile,sheet_name="Elenco verbali",skiprows=0)  #BDO	BDO_dum
#Replace nan values
df_work=df_work.fillna("")

"""RINOMINA le COLONNE non fossero con le intestazioni attese"""
df_work=df_work.rename(columns={"DipApprov":wcol_name_Dip})

#Elenca i valori distinti in una lista di Dip (Strutture aziendali)
wList_dip=df_work[wcol_name_Dip].unique()
#Per ciascuna struttura (Dip) presente nell'elenco 
for wnomDip in wList_dip:
#filtra il data frame per la struttura in esame
	df_work_filt= df_work.loc[df_work[wcol_name_Dip]==wnomDip]
#Inserisce in una lista le mail differenti riportate per la struttura
	wList_mail= df_work_filt[wcol_name_mail].unique()
#Inserisce in una lista le mail dei PMO riportate per la struttura
	wList_pmo_mail= df_work_filt["E-mail_pmo"].unique()
#Filtra nel dataframe le colonne che dovrenno essere riportate nella tabella HTML (in questo caso per posizione della colonna)
	df_work_filtcol=df_work_filt.iloc[:,wColfromExcel]
#Converte il dataframe in formato html
	wmail_table=df_work_filtcol.to_html(index=False) #la keyword index=False per non riportare l'indice nella prima colonna 
#Allineare a destra i numeri
	for n in range(0, 10):
		wmail_table = wmail_table.replace(f'<td>{n}', f'<td align="right">{n}')	
#	wmail_table=re.sub(r"<td>((\d)+?(.(\d)+))", r"<td class='my_class'>\1", df_work_filtcol.to_html())
#Converte le liste in elenco separato da ;
	wToMailaddress=';'.join([str(s) for s in wList_mail])
	wCcMailaddress=';'.join([str(s) for s in wList_pmo_mail])
#genera il testo della mail (inizio + taabelle html + fine)
	wmail_txt=wmail_txt_inz+wmail_table+wmail_txt_fin

#Additional cc for specific Department
	if wEnv=="FCan":
		wAddCcMail=";pmo-report@ariaspa.it" #;betsabea.rampoldi@ariaspa.it
	elif wEnv=="BRam":
		wAddCcMail=";pmo-report@ariaspa.it" #;betsabea.rampoldi@ariaspa.it
	elif wEnv=="EFil":
		wAddCcMail=";fabio.caneri@ariaspa.it;pmo-report@ariaspa.it"

	if wnomDip == "DCL": wAddCcMail=wAddCcMail#+";francesco.frigerio@ariaspa.it"
#Solo per la prima selezione stampa html di prova
	if wfirstloop:
		print("--- Primo Testo html---")
		print("Oggetto mail: ", wmail_oggetto+" - "+wnomDip)
		print("To_Dest: ",wToMailaddress)
		print("Cc_Dest: ",wCcMailaddress+wAddCcMail)
		print("---Testo html ---")
		print(html2text.html2text(wmail_txt))
		print("---fine testo html ---")
		wfirstloop=False

#inserisce nel dizionario i valori estratti
	wmail_dict={"Dip":wnomDip,"TestoMail":wmail_txt,"Oggetto":wmail_oggetto+" - "+wnomDip,"To_Dest":wToMailaddress,"Cc_Dest":wCcMailaddress+wAddCcMail}
#Aggiunge nell'ultima posizione del dataframe che contiene le mail quanto riportato nella struttura dizionario	
	df_tabmail.loc[len(df_tabmail)]=wmail_dict


#  **** Scrittura file excel con tabella ****
#Estrae le dimensioni del dataframe da scrivere in excel per la generazione della tabella	
(max_row, max_col) = df_tabmail.shape
#File excel da utilizzare per l'output
wrk_wrtName= wrk_dir_out+"\\"+"Mail_tosend.xlsx"

## Estrae le intestazioni di colonna per la creazione della tabella
column_settings = [{"header": column} for column in df_tabmail]

#Predispone il foglio di lavoro per output
workbook=xlsxwriter.Workbook(wrk_wrtName)

formatdatelong = workbook.add_format({'num_format': 'd mmm yyyy  hh:mm AM/PM'})
formatdate = workbook.add_format({'num_format': 'dd/mm/yy'})
# Widen column A for extra visibility.

#Aggiunge nel file Excel
worksheet1=workbook.add_worksheet()
#Crea la tabella associata
worksheet1.add_table(0, 0, max_row, max_col - 1, {"columns": column_settings,"data":df_tabmail.values.tolist(),"name": wNamTable})
#Scrive data e ora per visualizzare l'aggiornamento
worksheet1.write(0, max_col+2, "Ultimo aggiornamento")
worksheet1.write(1, max_col+2, datetime.now(),formatdatelong)
worksheet1.set_column(1, max_col-1, 25)
worksheet1.set_column(max_col+2, max_col+2, 20)
workbook.close()
