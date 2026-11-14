# STDlib
import logging, csv, re
from pathlib import Path

# My own modules
from .errors import sheetError
from .base import new_column_action, WriteOnlySpreadsheet, ReadOnlySpreadsheet, default
from .mytyping import * # For type checking tools
from .writer import WriterBase

log = logging.getLogger() # Set from main script if fancy logging is required.
# These should go to tools (and probably replaced with better tools by someone else)
    
class CSVFile: 
    FORMAT_NAME = "CSV"
    SUFFIXES = [".csv"]
    BOLDFONT = None

class WriterCSV(WriterBase, CSVFile): 
    # TODO: Convert to use the DictWriter class (needs data-pre-work to handle duplicate 'keys')
    """ """
    def __init__(self:Self, filename: str|Path, **kwargs:dict ):
        """Accept the same keyword arguments as pathlib.Path.open()."""
        WriterBase.__init__(self,filename)
        CSVFile.__init__(self)
        if self.SUFFIXES and Path(filename).suffix.lower() not in self.SUFFIXES: 
            raise sheetError(f"{self.FORMAT_NAME} inputput file name must end in {self.SUFFIXES[0]}. {self.fp} fails")
        # Need to store these for _read_header()
        self.dialect = kwargs.get("dialect","excel")

    def _format_save(self, fp:Path):
        assert self.file, "Should not save if file has not been opened yet!"
        if self.file: self.file.flush() # Write data to file immediately

    def _format_is_empty(self:Self):
        if not self.exists: return True
        return self.fp.is_file() and ( self.fp.stat().st_size == 0 )
    # ------- ABSTRACT METHOD IMPLEMENTATIONS -----------
    def _read_header(self:Self) -> tuple[str,...]:
        row = []
        with self.fp.open("r", 
                           encoding = self.encoding,
                           newline = self.newline,
                           errors = self.errors
                           ) as file:
            rdr = csv.DictReader(file, dialect = self.dialect)
            fields = rdr.fieldnames
            if not fields: return tuple()
        return tuple(fields)

    def open(self:Self,mode,
             fieldnames:Sequence[str] = tuple(),
             *args,
             **kwargs:dict) -> None:
        """Opens a CSV file for writing. Passes all arguments to csv.DictWriter()."""
        self.args = args # Keep a copy in case we need to reopen
        self.kwargs = kwargs # Keep a copy in case we need to reopen
        # If file exists, use first row as columns
        if not self.is_empty:
           # Get column data from existing file
           self.fields = self._read_header() + tuple(fieldnames)
        else:
            self.fields = fieldnames
        self.file = self.fp.open(mode, # Append mode
                                 *args,
                                 encoding = self.encoding,
                                 newline = self.newline,
                                 errors = self.errors                                 
                                 )
        self.writer = csv.DictWriter(self.file,*args,self.fields, **kwargs)
        self._opened = True

    def addrow(self:Self,
               datarowdict: dict[str,object],
               on_new_field:str = new_column_action.NEWFILE
               ) -> None:
        WriterBase.addrow(self,datarowdict, on_new_field) # Default behavior
        self.save() # CSV saving is fast: we can save after each row

    def on_insert_columns(self):
        raise NotImplementedError("WriterCSV() does not yet support column insertion.")

    def prep_outrow(self, row):
        #Overrides writerbase version
        return row
    
    def format_add_row(self, outrow):
        self.writer.writerow(  outrow )

    def writeheader(self:Self) -> None:
        # TODO: ERROR STATE HANDLING
        if not self.is_empty: pass # Do not write a header if file already has material in it
        self._saved = False
        # Can update  self.writer.fieldnames as long as header has not yet been written.
        self.writer.writeheader()
        assert self.file, "Should not save if file has not been opened yet!"
        self.save() # Write data to file immediately
        self._saved = True


class ReaderCSV(ReadOnlySpreadsheet, CSVFile): 
    # ------- ABSTRACT METHOD IMPLEMENTATIONS -----------
    def __init__(self:Self, filename: str|Path, **kwargs:dict): 
        ReadOnlySpreadsheet.__init__(self,filename, **kwargs)
        CSVFile.__init__(self)
        self.encoding = kwargs.get("encoding",default.ENCODING)
        self.newline = kwargs.get("newline",default.NEWLINE)
        self.errors = kwargs.get("errors",default.ENCODING_ERRORS)

    def open(self:Self, *args, **kwargs) -> None:
        """Opens a CSV file for reading. Passes all arguments to csv.reader()."""
        if not self.exists:
            raise sheetError(f"Could not access file {self.fp}.")
        self.file = self.fp.open("r",
                           encoding = self.encoding,
                           newline = self.newline,
                           errors = self.errors
                                    )
        self.reader = csv.reader(self.file, *args, **kwargs)
        self.fields = next(self) # Base class property uses this
        self._opened = True

    def __next__(self:Self):
        return next(self.reader)

