# STDlib
import logging, csv, re
from pathlib import Path

# Non-stdlib
from odfdo import Document, Row, Table

# My own modules
from errors import sheetError
from base import new_column_action, WriteOnlySpreadsheet, ReadOnlySpreadsheet
from mytyping import * # For type checking tools

def unique(s: Sequence) ->tuple[str,...] :return tuple(dict.fromkeys(s)) 
#BOLDFONT = Font(bold=True)
SUFFIXES = [".ods", "ods"]
FORMAT_NAME = "ODS"

log = logging.getLogger() # Set from main script if fancy logging is required.

class ODSFile():    
    def __init__(self:Self):
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
        self.wb = Document(filepath)
        if not self.wb:
            raise sheetError(f"Could not open file {filepath}")
        self.table = self.wb.body.tables[0]

    def format_create(self:Self):
        self.wb = Document("spreadsheet")
        self.wb.body.clear()
        self.table = Table(name="Table1")
        self.wb.body.append(self.table)

    def format_map_fields_to_columns(self:Self, fields:type_fields):
        return { k:v for k,v in zip(fields, range(len(fields))) } 

    def format_get_row_values(self:Self, nrow:int) -> tuple:
        assert self.table, "Self.table should exist when we ask for row values."
        return tuple( self.table.get_row(nrow).get_values() )

    def format_set_row_values(self:Self, nrow:int, values:Sequence) -> None:
        assert self.table, "Self.table should exist when we ask for row values."
        row = self.table.get_row(nrow)
        row.set_values(values)
        self.table.set_row(nrow, row)

    def format_save(self:Self, filepath:Path) -> None:
        assert self.wb, "Workbook should exist when we try to save."
        self.wb.save(filepath)
    
    def format_add_row(self:Self, values: Sequence):
        assert self.table, "Self.table should exist when we try to add a row."
        row = Row()
        for n in range(len(values)):
            row.set_value(n, values[n])
        self.table.append(row)


class WriteXLSX(WriteOnlySpreadsheet, ODSFile): 
    def __init__(self:Self, filename: str|Path): 
        WriteOnlySpreadsheet.__init__(self,filename)
        ODSFile.__init__(self) # This convenience class keeping ODS-specific loaders.
        self.field2col = {} # Maps column name to column number
        if self.fp.suffix.lower() not in SUFFIXES: 
            raise sheetError(f"{FORMAT_NAME} output file name must end in {SUFFIXES[0]}. {self.fp} fails")
    # ------- ABSTRACT METHOD IMPLEMENTATIONS -----------
    def open(self:Self,
             fieldnames:Sequence[str] = tuple()
             ) -> None:
        # If file exists, use first row as columns
        if not self.is_empty:
            self.format_open(self.fp)
            self.fields = self.format_read_header(self.fp) + tuple(fieldnames)
        else:
            self.format_create()
            self.fields = fieldnames
        # Map field name to column number (0-based indexing as we do not use this numbering directly in Excel)
        self.field2col = self.format_map_fields_to_columns(self.fields)
        self._opened = True

    def save(self:Self) -> None:
        self._saved = False
        self.format_save(self.fp)
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
        diff = self.lacks_columns( datarowdict.keys() )
        if diff: 
            log.warning(f"Missing column(s): {", ".join(diff)}.")
            if on_new_field == new_column_action.INSERT:
                log.info(f"Inserting new columns: {", ".join(diff)} into {self.filename}.")
                self.fields = self.fields + tuple(diff)
                self.field2col = { k:v for k,v in zip(self.fields, range(len(self.fields))) } 
                # Need to rewrite header line
                self.writeheader()
            else:
                sheetError(f"New column name, but no given way of adding columns specified.")
        outrow = [""]*len(self.fields)
        for colname,value in datarowdict.items():
            pos = self.field2col[colname]
            outrow[pos] = str(value) if value else ""
        row = Row()
        self.format_add_row(outrow)
        self.save() # Saving after each insert if quite heavy

    def writeheader(self:Self) -> None:
        assert self.wb, "self.wb sould exist when we try to writeheader()."
        if self.is_empty: # E,pty file, we need to add a row 
            self.format_add_row(self.fields)
        else: # Update existing header line
            self.format_set_row_values(0, self.fields)
        self.save() 

    def close(self, save:bool = True) -> None:
        if save: self.save()
        if not self._saved: log.warning(f"Closing file {self.filename} without saving first.")
        if self.wb: self.wb = None
        self._opened = False

class rowIterator:
    """Startrow is 0-indexed (first row = row 0)"""
    def __init__(self:Self, table, func:Callable, startrow:int = 1):
        self.table = table
        self.func = func
        self.startrow = 1 
        self.maxlen = int(table.size[1]) - 1
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

class ReadODS(ReadOnlySpreadsheet, ODSFile):
    def __init__(self:Self, filename: str|Path): 
        super().__init__(filename)
        self._reset_variables()
        if self.fp.suffix.lower() not in SUFFIXES: 
            raise sheetError(f"{FORMAT_NAME} input file name must end in {SUFFIXES[0]}. {self.fp} fails")

    def _reset_variables(self:Self):
        self._opened = False
        self.wb = None
        self.table = None
        self.fields = tuple()
        self.field2col = {  }  # Mapping field name to column number (0-indexed) Set by ODSFile open/create

    # ------- ABSTRACT METHOD IMPLEMENTATIONS -----------
    def open(self:Self) -> None:
        self._opened = False
        if not self.exists:
            raise sheetError(f"Coulf not access file {self.fp}.")
        self.format_open(self.fp)
        self.iter = iter(rowIterator(self.table, self.format_get_row_values))        
        self.fields = self.format_get_row_values(0)       
        self._opened = True

    def next(self:Self) -> Iterator[tuple]:
        return self.iter
        
    def close(self:Self) -> None:
        if self.wb: self.wb = None
        self._reset_variables()
        
    @property
    def fields(self:Self)-> tuple[str,...]: return tuple(self._fields)

    @fields.setter
    def fields(self:Self, val:Sequence) -> None:
        self._fields = tuple(unique(val))

if __name__ == '__main__':
    testfile = "test.ods"
    t = WriteXLSX( testfile )
    t.open(["col_foo", "col_bar"])
    t.add_line_from_dict({"col_foo":"test1", "col_bar":"test2"}, on_new_key=new_column_action.INSERT)
    t.add_line_from_dict({"col_foo":"testA", "col_bar":"testB"}, on_new_key=new_column_action.INSERT)
    t.add_line_from_dict({"col_foo":"testB", "col_bar":"testB", "col_baz":"testB"}, on_new_key=new_column_action.INSERT)
    t.add_line_from_dict({"col_foo":"testB",  "col_baz":"testB"}, on_new_key=new_column_action.INSERT)
    t.save()
    t.close()
    t = ReadODS( testfile )
    t.open()
    print("Fields = ", t.fields)
    for d in t.next():
        print(tuple(d))
    t.close()
