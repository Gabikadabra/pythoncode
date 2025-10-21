#Util for logging 
"""To use in different modules includind following code"""
""" 
** Begin **
from Util_logging import logger
wPythProc="read_Interni"
import timeit  #funzione timeit per valutare durata estrazione
tempoInzInt=timeit.default_timer()

** END **
df_RisInt.memory_usage() #Find memory usage
df_RisInt.memory_usage(deep=True) #Find each column usage
df_RisInt.memory_usage(deep=True).sum() #Find total usage
tempoFinInt = timeit.default_timer()

print("Durata Attivita *** ",tempoFinInt-tempoInzInt)
logger.info('end reading ' + str(df_RisInt.shape[0])+" Durata processo: " + wPythProc + " "+ str(round(tempoFinInt-tempoInzInt,3)))

logger.info('end reading ' + str(df_RisInt.shape[0])+" Durata processo: " + wPythProc +" - dim: "+str(df_RisInt.memory_usage().sum())+" Durata processo: "  + str(round(tempoFinInt-tempoInzInt,3))+ " Dataframe dimensione " + str(df_RisInt.memory_usage(deep=True)) )

if __name__ == '__main__':
    logger.info("This is an info message")
"""    
import logging

logger = logging
#logger= logging.getLogger( __name__ )
#logging.basicConfig(filename='example.log', encoding='utf-8', level=logging.DEBUG)
logger.basicConfig(filename='PMPRept.log',format='%(asctime)s - %(message)s', level=logging.INFO)