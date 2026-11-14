from pathlib import Path
# Expose main classes
from .writer import WriterXLSX, WriterODS
from .reader import ReaderXLSX, ReaderODS
from .jkcsv import WriterCSV, ReaderCSV
from .base import new_column_action

#convenience 
suffix2writer = {".xlsx":WriterXLSX,
                 ".ods":WriterODS,
                 ".csv":WriterCSV,
                 }
 

suffix2reader = {".xlsx":ReaderXLSX,
                 ".ods":ReaderODS,
                 ".csv":ReaderCSV,
                 }

def Reader(fp, *args, **kwargs):
	fp = Path(fp)
	return suffix2reader[fp.suffix](fp, *args, **kwargs)

def Writer(fp, *args, **kwargs):
	fp = Path(fp)
	return suffix2writer[fp.suffix](fp, *args, **kwargs)