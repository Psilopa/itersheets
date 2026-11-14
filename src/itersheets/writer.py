# STDlib
import logging
from pathlib import Path
from datetime import datetime

# My own modules
from .jkods import ODSFile
from .jkxlsx import XLSXFile
from .errors import sheetError
from .base import new_column_action, WriteOnlySpreadsheet, Spreadsheet
#from .mytyping import * # For type checking tools

log = logging.getLogger() # Set from main script if fancy logging is required.

def sanitize_filename(text:str) -> str:
    return text
    dirty = "".join(["\\[/\\?%*:|\"<>0x7F0x00-0x1F]", chr(0x1F) * 15])
    return re.sub(text,"-",dirty) # Replace all chars in dirty with -.

# -------------- Shared code for XLSX and ODS implementation --------------
class WriterBase(WriteOnlySpreadsheet):

    def __init__(self:Self, filename: str|Path): 
        super().__init__(filename)
        self.field2col = {} # Maps column name to column number
        if self.SUFFIXES and Path(filename).suffix.lower() not in self.SUFFIXES: 
            raise sheetError(f"{self.FORMAT_NAME} inputput file name must end in {self.SUFFIXES[0]}. {self.fp} fails")

    @Spreadsheet.is_empty.getter
    #We need to reimplement this as Sheet Writers do not always exist as a file right away
    def is_empty(self:Self) -> bool:
        return self._format_is_empty()
        
    # BASE ABSTRACT METHOD IMPLEMENTATIONS
    def open(self:Self,mode,
             fieldnames:Sequence[str] = tuple(),
             *args,
             **kwargs:dict) -> None:
        self._reset()
        print("***** IN WRITER:OPEN*****")
        self.args = args # Keep a copy in case we need to reopen
        self.kwargs = kwargs # Keep a copy in case we need to reopen
            
        # If file exists, use first row as columns
        if self.exists:
            print("EXISTS")
            print("FILE PATH = ", self.fp)
            self.format_open(self.fp)
            self.fields = self.format_read_header(self.fp) + tuple(fieldnames)
        else:
            print("NOT EXISTS")
            self.format_create()
            self.fields = fieldnames
        # Map field name to column number (0-based indexing as we do not use this numbering directly in Excel)
        self.field2col = self.format_map_fields_to_columns(self.fields)        
        self._opened = True

    def _reset(self:Self):
        self._opened = False
        self.fields = tuple()
        self.field2col = {} # Maps column name to column number

    def save(self:Self) -> None:
        if self._saved: return # Nothing to do
        self._saved = False
        self._format_save(self.fp)
        self._saved = True

    def addrow(self:Self,
               datarowdict: dict[str,object],
               on_new_field:str
               ) -> None:
        self._saved = False
        if not self.is_open:
            raise sheetError(f"Trying to write into {self.fp} after closing the file.")        
        if self.is_empty:
            self.writeheader()
        # Remove any data with field name = None
        # NOTE: THIS FAILS IF WE ALLOW EMPTY HEADERS IN THE FUTURE
        if None in datarowdict: del datarowdict[None]

        diff = self.lacks_columns( datarowdict.keys() )
        if diff:
            if len(self.fields) == 0: # We had no headers previosly (new file)
                self.fields = datarowdict.keys()
            else:
                log.warning(f"Missing column(s): {", ".join(diff)}.")
                self.fields = self.fields + tuple(diff)
            if on_new_field == new_column_action.NEWFILE:
                self.on_newfile()
                # Force header writing
                self.writeheader()
            elif on_new_field == new_column_action.INSERT:
                log.info(f"Inserting new columns: {", ".join(diff)} into {self.filename}.")
                self.on_insert_columns()
            else:                
                raise sheetError(f"New column name, but no given way of adding columns specified. WriterCSV supports only swapping to a new file for now.")            
        outrow = self.prep_outrow(datarowdict)
        self.format_add_row(outrow)
#        self.save() # Saving after each insert is quite heavy!

    def on_insert_columns(self):
        self.field2col = { k:v for k,v in zip(self.fields, range(len(self.fields))) } 
        self.writeheader()

    def on_newfile(self):
        if not self.is_empty: # Not an empty file, need to create a backup
            oldname = self.fp.stem
            timestamp = datetime.now().strftime("%Y-%m-%d-%p%I-%M-%S")
            newname = Path(oldname + "-" + timestamp).with_suffix(self.fp.suffix)
            newname = Path( sanitize_filename(newname) )
            log.warning(f"Moving old file to {newname}")
            self.move(newname, overwrite = True)
            newfields = self.fields # Grab before closing the old file
            self.close()
        log.warning(f"Entering new data to {self.filename}.")
        # Reopen original filename for data
        self.open("w", fieldnames = newfields, *self.args, **self.kwargs) # Reopen as a new file

    def prep_outrow(self, datarowdict): 
        outrow = [""]*len(self.fields)
        for colname,value in datarowdict.items():
            pos = self.field2col[colname]
            outrow[pos] = str(value) if value else ""
        return outrow

    def writeheader(self:Self) -> None:
        self._saved = False
        if self.is_empty: # Empty file, we need to add a row 
            self.format_add_row(self.fields)
        else: # Update existing header line
            self.format_set_row_values(0, self.fields)
        if self.BOLDFONT: self.format_set_row_font(0, self.BOLDFONT)
        # Do not save just for a header, wait for a data row

    def close(self, save:bool = True) -> None:
        if save: self.save()
        if not self._saved: log.warning(f"Closing file {self.filename} without saving first.")
        self._reset()

#------------------ EXCEL SPECIFIC STUFF -------------
class WriterXLSX(WriterBase, XLSXFile): 
    def __init__(self:Self, filename: str|Path): 
        WriterBase.__init__(self,filename)
        XLSXFile.__init__(self)
#------------------ ODS SPECIFIC STUFF -------------
class WriterODS(WriterBase, ODSFile): 
    def __init__(self:Self, filename: str|Path): 
        WriterBase.__init__(self,filename)
        ODSFile.__init__(self)
