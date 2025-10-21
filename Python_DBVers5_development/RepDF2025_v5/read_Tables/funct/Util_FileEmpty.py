#Controlla se file non è vuoto
import os

def file_is_empty(path):
    return os.stat(path).st_size==0