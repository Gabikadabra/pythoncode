#Write Excel files con il contenuto del dataframe

import os
import pandas as pd
from Util_leggeDir import dir_dict
wrk_dir_inp=dir_dict["inpDir"]
from read_map import df_MAP_byMAP,df_MAP_bytask,df_pagam_tot
from elab_Oda_Periodo import df_ODA
df_pagam_tot.to_excel(wrk_dir_inp+"\\"+"ListamapFatt.xlsx", sheet_name="foglio1", index=False, startrow=0, startcol=0)
df_MAP_byMAP.to_excel(wrk_dir_inp+"\\"+"Listamap.xlsx", sheet_name="foglio1", index=False, startrow=0, startcol=0)
df_ODA.to_excel(wrk_dir_inp+"\\"+"ListaODA.xlsx", sheet_name="foglio1", index=False, startrow=0, startcol=0)

from elab_taskMAP import  df_MAP_bytask_fatt,df_prog_bytask

df_MAPtask=df_MAP_bytask_fatt.groupby("Task")["MAP_Cons"].sum().reset_index()
df_progtask=df_prog_bytask.groupby("Codice Task")[["Pianif_Corr","Consuntivo"]].sum().rename(columns={"ValoreMap":"MAP_cons","Codice Task":"Task"}).reset_index()
#df_MAP_bytask.to_excel(wrk_dir_inp+"\\"+"Listamaptask.xlsx", sheet_name="foglio1", index=False, startrow=0, startcol=0)
df_progtask=df_progtask.rename(columns={"Codice Task":"Task"})

df_progtask=pd.merge(df_progtask,df_MAPtask,on="Task",how="outer",indicator=True,suffixes=('', 'prog'))

df_progtask.to_excel(wrk_dir_inp+"\\"+"ListaTask.xlsx", sheet_name="foglio1", index=False, startrow=0, startcol=0)