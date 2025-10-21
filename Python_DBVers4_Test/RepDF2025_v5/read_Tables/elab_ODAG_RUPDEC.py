#Get DEC & RUP information for ODAG
import pandas as pd
#read tables
from get_ctrMgmt import df_ODAG_ctrMgmt,df_ODA_ctrMgmt
from read_ODAG import df_ODAGMacr, df_ODAG
#get DEC & RUP information
df_ODAG=pd.merge(df_ODAG,df_ODAG_ctrMgmt[["ODAGnbr","DEC","RUP"]],left_on="ODAGnbr",right_on="ODAGnbr", how="left",suffixes=('', '_bdo'))
df_ODAG=df_ODAG[(df_ODAG["RUP"].notnull())&(df_ODAG["RUP"].notnull())] # Eliminate ODAG without DEC AND RUP
#df_ODAG=df_ODAG.drop(["ODAGnbr"],axis=1)
df_ODAGMacr=pd.merge(df_ODAGMacr,df_ODAG_ctrMgmt[["ODAGnbr","DEC","RUP"]],left_on="ODAGnbr",right_on="ODAGnbr", how="left",suffixes=('', '_bdo'))
#df_ODAGMacr=df_ODAGMacr.drop(["ODAGnbr"],axis=1)
pass