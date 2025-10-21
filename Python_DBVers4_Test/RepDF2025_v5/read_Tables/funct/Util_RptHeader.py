import xlsxwriter
   #predispone funzioni di utilità
from .Util_ReportDate import GetRptDate,GetAlfb
		#import datetime   #importare libreria per la data del giorno
from datetime import timedelta,datetime
from read_rptDate import wRptDate
##2 from .Util_wDir import RepInit0	#Per la scrittura su Excel si usa la libreria xlsxwriter (limite che non si possono leggere file Excel e non si possono modificare con quella libreria)

def EstraeRepIntestaz(worksheet, list):
    # Create label
    worksheet.write(0, 0, "Intestazione Report")

    # Code here to loop and add list elements row by row

def RepTitoloData(fWorkbook, fWorksheet,fTitolo):
 
	#Estrae le dir di input (wdirInp)  e quelle di output (wdirOut)
##2	wdirInp,wdirOut=RepInit0()	
	# Imposta formato intestazione colonne (sfondo verde)
	BoldBlue_fmt=fWorkbook.add_format({'bold': True, 'font_size':12})
	#bckgrdGreen_fmt = fWorkbook.add_format()
	#bckgrdGreen_fmt.set_bg_color("#42f58d")
    #Estrae data ultimo aggiornamento
##2	wDatUpdate = GetRptDate(wdirInp)
#Predispone il formato delle date che verranno utilizzate per i report
	wDtFormat = "%d/%m/%Y"  #Esegue la procedura per ciascuna tipologia di output
	wDate_today = datetime.now().strftime("%d/%m/%Y")
    #Formato intestazione
	fWorksheet.write("A1", fTitolo, BoldBlue_fmt)
	fWorksheet.write("A2", "Prod.il")
	fWorksheet.write("B2", wDate_today )
	fWorksheet.write("A3","Estraz.del")
	fWorksheet.write("B3",wRptDate)
#Ambiente di test  


    #header1_format = workbook.add_format()
	#fWorksheet.write(1, 0, fTitolo, header1_format)


#def RepFormat(fWorkbook, fWorksheet,fColumn,fFormat):
def RepColumnFormat(fWorkbook, fWorksheet,fListCol,fFmt,fCol_fm=0,fCol_to=0):
			# Imposta Formato currency
	money_fmt = fWorkbook.add_format({'num_format': '#,##0.00'})
	#if fListCol==True:
	fColNbr	=GetAlfb(fListCol)  # converte la lista da alfab a num.colonna Excel
	#else:
	#	fColNbr=list(range(fCol_fm,fCol_to+1))
	#formato richiesto -> formattazione in base 
	if fFmt == "10":
		for colm in fColNbr: fWorksheet.set_column(colm, colm, 10)
	elif fFmt == "12":
		for colm in fColNbr: fWorksheet.set_column(colm, colm, 12)
	elif fFmt == "15":
		for colm in fColNbr: fWorksheet.set_column(colm, colm, 15)
	elif fFmt == "20":
		for colm in fColNbr: fWorksheet.set_column(colm, colm, 20)
	elif fFmt == "30":
		for colm in fColNbr: fWorksheet.set_column(colm, colm, 30)
	elif fFmt == "40":
		for colm in fColNbr: fWorksheet.set_column(colm, colm, 40)
	elif fFmt == "ccy":
		for colm in fColNbr: fWorksheet.set_column(colm, colm, 12, money_fmt)
# da utilizzare per valori che superano i 100.000.000,00
	elif fFmt == "ccyL":
		for colm in fColNbr: fWorksheet.set_column(colm, colm, 13, money_fmt)
	else:
		print("formato non trovato")

