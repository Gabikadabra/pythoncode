#Write Dataframe to excel
from Util_leggeDir import dir_dict
def write_templShname(f_reptname,f_lib_dframe,f_wrk_dframe):

    package = f_lib_dframe
    name = f_wrk_dframe
    # "from  import" using variables
    imported = getattr(__import__(package, fromlist=[name]), name)

#check if dataframe sheet exist
#    from openpyxl import load_workbook
    import openpyxl
    wrk_CsvDir=dir_dict["csvDir"]
    wFmtFile="RptFormati"
# Load existing excel file into a openpyxl Workbook object
    wBook = openpyxl.load_workbook(wrk_CsvDir+"\\"+wFmtFile+".xlsx")
# If sheet 'testSheet' does not exist yet, then add it in the openpyxl Workbook object
    if not f_reptname in wBook.sheetnames:
        wBook.create_sheet(f_reptname)
        index_list = imported.columns.values.tolist()
        #imported.index.tolist()
        wSheet = wBook[f_reptname]
        wSheet.cell(1,1).value ="ColumnName"
        wSheet.cell(1,2).value ="ColumnAlias"
        wSheet.cell(1,3).value ="Riepilogo"
        i=2
        for item in index_list:
            #wSheet.append(item)
            wSheet.cell(i,1).value = item
            i=i+1

# Save the openpyxl Workbook object to file
    wBook.save(wrk_CsvDir+"\\"+wFmtFile+".xlsx")    


    #    df_new.to_excel(writer, sheet_name='New Sheet', index=False)
