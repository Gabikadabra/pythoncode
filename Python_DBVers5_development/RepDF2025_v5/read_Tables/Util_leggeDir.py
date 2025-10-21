#Legge le directory di lavoro
#Funzione per estrarre le directory di lavoro
#Inizializza variabili principali
import os
#2)	Gestione DIRECTORY CORRENTE (contestualizzazione cartella python sulla base dell’utente)
#da inserire -> selezione directory sulla base dell'utente che utilizza Python
w_environm_inp="PROD"#"Testfilter"#"Prod" #"testpmoshared" "prod" "Testfilter"
w_environm_out="Test"
current_dir=os.getcwd()
    #condizione if
if os.getenv("username") == "fcaneri":     #se utente è Fabio Caneri
    wrk_home_dir= r"\Users\fcaneri\OneDrive - ARIA S.p.A" #inserisce in variabile wrk_home_dir la directory di ri
    wrk_read_dir= r"\Users\fcaneri\OneDrive - ARIA S.p.A\Documenti - Share PMO Data" #inserisce in variabile wrk_home_dir la directory di ri
    wrk_ctrMgmt_dir= r"\Users\fcaneri\OneDrive - ARIA S.p.A\General - Monitoraggio contratti" #inserisce in variabile wrk_home_dir la directory di ri
    wrk_progoff_dir= r"\Users\fcaneri\OneDrive - ARIA S.p.A\Materiale incarichi\2024" #inserisce in variabile wrk_home_dir la directory di ri
    wEnv="FCan"
elif os.getenv("username") == "brampoldi":
    wrk_home_dir= r"\Users\brampoldi\OneDrive - ARIA S.p.A" #inserisce in variabile wrk_home_dir la directory di ri
    wrk_read_dir= r"\Users\brampoldi\OneDrive - ARIA S.p.A\Documenti - Share PMO Data" #inserisce in variabile wrk_home_dir la directory di ri
    wrk_ctrMgmt_dir= r"\Users\brampoldi\OneDrive - ARIA S.p.A\General - Monitoraggio contratti" #inserisce in variabile wrk_home_dir la directory di ri
    wrk_progoff_dir= r"\Users\brampoldi\OneDrive - ARIA S.p.A\Materiale incarichi\2024" #inserisce in variabile wrk_home_dir la directory di ri
    wEnv="BRam"
elif os.getenv("username") == "filippone.elisabetta" or os.getenv("username") == "utente":
    wrk_home_dir= r"\Users\utente\OneDrive - ARIA S.p.A"
    wrk_read_dir= r"\Users\utente\OneDrive - ARIA S.p.A\Documenti - Share PMO Data"
    wrk_ctrMgmt_dir= r"\Users\utente\OneDrive - ARIA S.p.A\General - Monitoraggio contratti" #inserisce in variabile wrk_home_dir la directory di ri
    wrk_progoff_dir= r"\Users\utente\OneDrive - ARIA S.p.A\Materiale incarichi" #inserisce in variabile wrk_home_dir la directory di ri

    wEnv="EFil" #sigla utenza che elabora lo script
    #condizione alternativa \Users\fcaneri\OneDrive - ARIA S.p.A\PMO Condiviso\Report test\Python\Develop\ReportPM

elif os.getenv("username") == "filippone.elisabetta" or os.getenv("username") == "BC-UNKNOWN":
    wrk_home_dir= r"\Users\BC-UNKNOWN\Desktop\OneDriveAria\OneDrive - ARIA S.p.A"
    wrk_read_dir= r"\Users\BC-UNKNOWN\Desktop\OneDriveAria\OneDrive - ARIA S.p.A\Documenti - Share PMO Data"
    wrk_ctrMgmt_dir= r"\Users\BC-UNKNOWN\Desktop\OneDriveAria\OneDrive - ARIA S.p.A\General - Monitoraggio contratti" #inserisce in variabile wrk_home_dir la directory di ri
    wrk_progoff_dir= r"\Users\BC-UNKNOWN\Desktop\OneDriveAria\OneDrive - ARIA S.p.A\Materiale incarichi" #inserisce in variabile wrk_home_dir la directory di ri

    wEnv="EFil" #sigla utenza che elabora lo script
else:
    print("user non trovato")
    exit()

#os.listdir(wrk_read_dir)
#os.listdir(wrk_home_dir)
wFolderZip_path = wrk_home_dir+r"\ImportDati\Staging\compress"
wFolderInpProd_path = wrk_read_dir+r"\ImportDati"
wFolderIF_path =wrk_read_dir+r"\Nuovo Processo BdO\File di appoggio"
#if "Develop" in current_dir:

if w_environm_inp.lower()=="test":
    wFolderXlsx_path = wrk_home_dir+r"\PMO Condiviso\Report test\Staging\xlsx"
    wFolderStag_path = wrk_home_dir+r"\PMO Condiviso\Report test\Staging"
    wFolderXls_path = wrk_home_dir+r"\PMO Condiviso\Report test\Staging\xls"
    wFolderTmp_path = wrk_home_dir+r"\PMO Condiviso\Report test\Staging\tmp"
    wFolderCSV_path =wrk_home_dir+r"\PMO Condiviso\Report test\csv3"
    wFolderInp_path = wrk_home_dir+r"\PMO Condiviso\ImportDati"
elif w_environm_inp.lower()=="testpmoshared":
    wFolderXlsx_path = wrk_home_dir+r"\PMO Shared\Staging\xlsx"
    wFolderStag_path = wrk_home_dir+r"\PMO Shared\Staging"
    wFolderXls_path = wrk_home_dir+r"\PMO Shared\Staging\xls"
    wFolderTmp_path = wrk_home_dir+r"\PMO Shared\Staging\tmp"
    wFolderCSV_path =wrk_home_dir+r"\PMO Condiviso\Report test\csv3"
    wFolderInp_path = wrk_home_dir+r"\PMO Shared\ImportDati"
else:
    #wFolderIn_path = wrk_home_dir+r"\ImportDati\Staging\compress"
    #wFolderOut_path = wrk_home_dir+r"\ImportDati\Staging\xlsx"
    #wFolderXls_path = wrk_home_dir+r"\ImportDati\Staging\xls"
    #wFolderTmp_path = wrk_home_dir+r"\ImportDati\Staging\tmp"
    wFolderXlsx_path = wrk_read_dir+r"\ImportDati\Staging\xlsx"
    wFolderXls_path = wrk_read_dir+r"\ImportDati\Staging\xls"
    wFolderTmp_path = wrk_read_dir+r"\ImportDati\Staging\tmp"
    wFolderStag_path = wrk_read_dir+r"\ImportDati\Staging"

    wFolderCSV_path =wrk_home_dir+r"\PMO Condiviso\Report Produzione\csv3"
    wFolderInp_path = wrk_read_dir+r"\ImportDati"

if w_environm_out.lower()=="test":
    wFolderRpt_path =wrk_home_dir+r"\PMO Condiviso\Report test\Report\OutputProcPython"
else:
    wFolderRpt_path =wrk_home_dir+r"\PMO Condiviso\Report Produzione\Report\OutputProcPython"

# Posiziona nella cartella di lavoro
wrk_dir = wrk_home_dir+r"\PMO Condiviso\Report test\Python"
os.chdir(wrk_dir)
new_dir = os.getcwd()

# Create empty dictionary
dir_dict = {}
#add new dictionary item
dir_dict["IFprcDir"] = wFolderIF_path
dir_dict["zipDir"] = wFolderZip_path
dir_dict["xlsxDir"] = wFolderXlsx_path
dir_dict["xlsDir"] = wFolderXls_path 
dir_dict["stagDir"] = wFolderStag_path 
dir_dict["tmpDir"] = wFolderTmp_path 
dir_dict["rptDir"] =     wFolderRpt_path
dir_dict["csvDir"] =     wFolderCSV_path 
dir_dict["inpDir"] = wFolderInp_path
dir_dict["inpDirProd"]=wFolderInpProd_path
dir_dict["ctrMgmt"] = wrk_ctrMgmt_dir
dir_dict["ProgOff"] =  wrk_progoff_dir+r"\2024"
dir_dict["ShrPtDir"]=wrk_read_dir
dir_dict["homeDir"]=wrk_home_dir
dir_dict["inpEnv"]=w_environm_inp
dir_dict["outEnv"]=w_environm_out