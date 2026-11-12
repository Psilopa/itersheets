# STDlib
import logging
from pathlib import Path

# Non-stdlib
from odfdo import Document, Row, Table

# My own modules
from .errors import sheetError
from .base import ReadOnlySpreadsheet
#from .mytyping import * # For type checking tools

def unique(s: Sequence) ->tuple[str,...] :return tuple(dict.fromkeys(s)) 

log = logging.getLogger() # Set from main script if fancy logging is required.

class ODSFile():    
    FORMAT_NAME = "ODS"
    SUFFIXES = [".ods"]
    BOLDFONT = None #Font(bold=True)
    def __init__(self:Self):
        self.formar_reset()
        if self.SUFFIXES and self.fp.suffix.lower() not in self.SUFFIXES: 
            raise sheetError(f"{self.FORMAT_NAME} inputput file name must end in {self.SUFFIXES[0]}. {self.fp} fails")

    def formar_reset(self:Self):
        self.wb = None
        self.table = None
        
    def format_read_header(self:Self, filepath:Path) -> tuple:
        tmpwb = Document(filepath)
        assert tmpwb, "wb should exist when we try to read a row in _read_header_values()."
        tmptable = tmpwb.body.tables[0]
        maxlen = int(tmptable.size[0])
        rv = unique([tmptable.get_row(0).get_value(x) for x in range(0, maxlen)])
#        rv = tuple(x for x in rv if x!=None) # Remove Nones
        tmpwb = None 
        return rv

    def format_open(self:Self, filepath: Path):
        if not filepath.exists():  # TODO: test for readable.
            raise sheetError(f"File {filepath} does not exist or is unavailable.")
        self.wb = Document(filepath)
        if not self.wb:
            raise sheetError(f"Could not open file {filepath}")
        self.table = self.wb.body.tables[0]

    def format_create(self:Self):
        self.wb = Document("spreadsheet")
        if not self.wb:
            raise sheetError(f"Could not create a new Excel file object.")
        self.wb.body.clear()
        self.table = Table(name="Table1")
        self.wb.body.append(self.table)

    def format_map_fields_to_columns(self:Self, fields:type_fields):
        return { k:v for k,v in zip(fields, range(len(fields))) } 

    def format_add_row(self:Self, values: Sequence):
        assert self.table, "Self.table should exist when we try to add a row."
        row = Row()
        for n in range(len(values)):
            row.set_value(n, values[n])
        self.table.append(row)

    def format_get_row_values(self:Self, nrow:int) -> tuple:
        assert self.table, "Self.table should exist when we ask for row values."
        return tuple( self.table.get_row(nrow).get_values() )

    def format_set_row_values(self:Self, nrow:int, values:Sequence) -> None:
        assert self.table, "Self.table should exist when we ask for row values."
        row = self.table.get_row(nrow)
        row.set_values(values)
        self.table.set_row(nrow, row)

    def format_set_row_font(self:Self, nrow:int, font):
        pass

    def format_save(self:Self, filepath:Path) -> None:
        # TODO: ERROR HANDLING ON BAD PATH
        assert self.wb, "Workbook should exist when we try to save."
        self.wb.save(filepath)

    def format_close(self:Self, filepath:Path|None=None) -> None:
        if self.wb and filepath:
            self.format_save(filepath)
        self.formar_reset()

    def value_iterator(self:Self):
        return iter( rowIterator( self.format_get_row_values,
                                  int(self.table.size[1]) - 1,
                                  startrow = 1
                                  )
                    )

# READING
class rowIterator:
    """Startrow is 0-indexed (first row = row 0)"""
    def __init__(self:Self, func:Callable, maxlen:int, startrow:int = 0):
        self.func = func
        self.startrow = startrow 
        self.maxlen = maxlen
    def __iter__(self:Self):
        self.nrow = self.startrow
        return self
    def __next__(self:Self):
        if self.nrow <= self.maxlen:
            rv = self.func(self.nrow)
            self.nrow += 1
            return rv
        else:
          raise StopIteration


        

