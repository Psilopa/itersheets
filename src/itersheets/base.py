# DEPENDS ON PYTHON 3.14? 
# STDlib
from abc import ABC, abstractmethod
from pathlib import Path
import logging, json, itertools

# My own modules
from .mytyping import * # For type checking tools
from .errors import sheetError
from .tools import unique

log = logging.getLogger() # Set from main script if fancy logging is required.

class default:
    ENCODING_ERRORS = "convert"
    ENCODING = "utf8"
    NEWLINE = ""

class new_column_action:
    NEWFILE = "newfile" # Only supported for CSV at the moment
    INSERT = "insert" # Only supported for XLS at the moment
#   EXCEPTION = "reject" # TODO- not currently supported. Raise a UnknownColumnError ?

def json2printdict(jsondata: str | dict[str,Any] ) -> dict[str,Any]:
    "Converts a string with JSON data or a corresponding dictionary with prettyprinted content."
    # TODO: this is not pasticularly good
    if isinstance(jsondata, str):  jsondata = json.loads(jsondata) 
    # Reformat a JSON reply to for Table output
    rv: dict[str,str] = {}
    for key in jsondata.keys():
        if isinstance(jsondata[key], list):
            rv[key] = list2printvalues( jsondata[key] )
        else: rv[key] = jsondata[key]
    return rv

def list2printvalues( mylist: Collection[dict] )-> str:
    "Prettyprints the values from a list of dictionaries."
    out = []
    for valuedict in mylist:
        out +=  valuedict.values()
    return "\n".join(out)


class Spreadsheet(ABC):
    def __init__(self:Self, filename: str|Path, **kwargs:dict ) : 
        self.encoding = kwargs.get("encoding",default.ENCODING)
        self.newline = kwargs.get("newline",default.NEWLINE)
        self.errors = kwargs.get("errors",default.ENCODING_ERRORS)
        self.fp = Path(filename)
        self.file = None
        self._opened = False
        self._fields = tuple()
        self._fieldset = set(self._fields)

    @property
    def filename(self:Self) -> Path: return self.fp

    @property
    def is_open(self:Self) -> bool: return self._opened

    @property
    def exists(self:Self) -> bool:
        if not self.fp or not self.fp.exists(): return False
        else: return True

    @property
    def is_empty(self:Self) -> bool:
        if not self.exists: return True
        return self.fp.is_file() and ( self.fp.stat().st_size == 0 )

    @property
    def fields(self:Self)-> tuple[str,...]: return self._fields

    @fields.setter
    def fields(self:Self, val) -> None:
        self._fields = unique( val )
        self._fieldset = set( self._fields )

    def close(self:Self) -> None:
        "Closes the associated file(s) and marks the Spreadsheet as not open."
        if self.file: self.file.close()
        self._opened = False

    def has_columns(self:Self, fields:Collection[str]) -> bool:
        "True is spreadsheet has a field for each value in fields."
        return self._fieldset.issuperset( set(fields) )

    def lacks_columns(self:Self, fields:Collection[str]) -> set[str]:
        "Returns a set of field names that are missing from the spreadsheet."
        return set(fields).difference( self._fieldset )

# ------------ WRITE ONLY SPREADSHEETS -------------    
class WriteOnlySpreadsheet(Spreadsheet):
    def __init__(self:Self, filename: str|Path, **kwargs:dict ):
        super().__init__(filename, **kwargs)
        self._saved = False

    @property
    def is_saved(self:Self) -> bool: return self._saved

    @abstractmethod        
    def open(self:Self,mode,
             fieldnames:Sequence[str] = tuple(),
             *args,
             **kwargs:dict) -> None:
        raise NotImplementedError

    # Override parent class methods in child classes
    def close(self:Self, save:bool=True) -> None:
        if save: self.save()
        if not self._saved: log.warning(f"Closing file {self.filename} without saving first.")
        super().close()

    @abstractmethod        
    def save(self:Self): 
        raise NotImplementedError
        
    @abstractmethod        
    def addrow(self:Self,
               datarowdict: dict[str,object],
               on_new_field:str 
               ) -> None:
        raise NotImplementedError

   
    def add_row_from_dict(self:Self,
                           datarowdict: dict[str,object],
                           on_new_key:str ) -> None:
        """Write a row based on a dictionary of column_header: value items.

        If the column_header is not included in the existing file structure, raise a ValueError."""
        self.addrow( json2printdict(datarowdict), on_new_key )
        
    def add_row_from_json(self:Self, datarowdict: dict[str,object],
                           on_new_key:str  ) -> None:
        self.addrow( json2printdict(datarowdict), on_new_key )

    def move(self:Self, newpath, overwrite=False):
        if not self.exists: return # Nothing to do.
        if newpath.exists() and not overwrite:
            raise sheetError(f"Will not overwrite {newpath}")
        self.fp.replace(newpath)
# -------------- READING ----------------
class ReadOnlySpreadsheet(Spreadsheet):
    def __init__(self:Self, filename: str|Path, **kwargs:dict) :
        super().__init__(filename, **kwargs)

    @abstractmethod        
    def open(self:Self, *args, **kwargs) -> None:
        self._opened = True

    # Iteration methods
    def __iter__(self):
        "Should set up a starting point for iteration"
        return self

    @abstractmethod        
    def __next__(self:Self):
        "Next row as a list"
        raise NotImplementedError

    def dictrows(self:Self):
        """An iterator returning rows as a dictionary of column_header: cell value pairs.

        Only available if row 1 has unique column names.
        """
        # itertools.zip_longest() pads the shorter iterable with Nones
        for row in self: yield {k:v for k,v in itertools.zip_longest(self.fields,row)}
                    
    def tuplerows(self:Self):
        "For consistency with dictrows, returns rows as tuples"
        for row in self: yield tuple(row)


