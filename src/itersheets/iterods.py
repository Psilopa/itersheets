# STDlib
import logging, csv, re
from pathlib import Path

# Non-stdlib
from odfdo import Document, Row, Table
#from openpyxl import Workbook, load_workbook
#from openpyxl.styles import Font

# My own modules
from errors import sheetError
from base import new_column_action, WriteOnlySpreadsheet, ReadOnlySpreadsheet
#from .mytyping import * # For type checking tools

def unique(s: Sequence) ->tuple[str,...] :return tuple(dict.fromkeys(s)) 
#BOLDFONT = Font(bold=True)
SUFFIXES = [".ods", "ods"]
FORMAT_NAME = "ODS"

log = logging.getLogger() # Set from main script if fancy logging is required.


class ODSFile():
    def _read_header(self:Self) -> tuple:
        wb = Document(self.fp)
        assert wb, "wb should exist when we try to read a row in _read_header()."
        table = wb.body.tables[0]
        maxlen = 50
        rv = unique([table.get_row(0).get_value(x) for x in range(0, maxlen)])
        rv = tuple(x for x in rv if x!=None) # Remove Nones
        wb = None 
        return rv

    def open_ods(self, fieldnames=tuple()):
        self.wb = Document(self.fp)
        self.table = self.wb.body.tables[0]
        if not self.table:
            raise sheetError(f"Could not find sheet {sheetname} in table {self.fp}")
        # Map field name to column number (0-based indexing as we do not use this numbering directly in Excel)
        self.fields = self._read_header() + tuple(fieldnames)
        self.field2col = { k:v for k,v in zip(self.fields, range(len(self.fields))) } 

    def create_ods(self, fieldnames=tuple()):
        self.wb = Document("spreadsheet")
        self.wb.body.clear()
        self.table = Table(name="Table1")
        self.wb.body.append(self.table)
        # Map field name to column number (0-based indexing as we do not use this numbering directly in Excel)
        self.fields = fieldnames
        self.field2col = { k:v for k,v in zip(self.fields, range(len(self.fields))) } 



class WriteXLSX(WriteOnlySpreadsheet, ODSFile): 
    def __init__(self:Self, filename: str|Path): 
        super().__init__(filename)
        self.wb = None
        self.field2col = {} # Maps column name to column number
        if self.fp.suffix.lower() not in SUFFIXES: 
            raise sheetError(f"{FORMAT_NAME} output file name must end in {SUFFIXES[0]}. {self.fp} fails")
    # ------- ABSTRACT METHOD IMPLEMENTATIONS -----------
    def open(self:Self,
             fieldnames:Sequence[str] = tuple()
             ) -> None:
        # If file exists, use first row as columns
        if not self.is_empty:
            self.open_ods(fieldnames)
        else:
            self.create_ods(fieldnames)
        self._opened = True

    def save(self:Self) -> None:
        self._saved = False
        self.wb.save(self.fp)
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
        self._addrow(outrow)
        self.save() # Saving after each insert if quite heavy

    def _addrow(self,values):
        row = Row()
        for n in range(len(values)):
            row.set_value(n, values[n])
        self.table.append(row)

    def writeheader(self:Self) -> None:
        # TODO: THIS SHOULD ALWAYS ACCESS LINE 1
        if not self.exists:
            self._addrow(self.fields)
            self.save() # Saving after each insert if quite heavy
        else:
            row = self.table.get_row(0)
            row.set_values(self.fields)
            self.table.set_row(0, row)
        # TODO: UPDATING ROW 0 IF IT DID ALREADY EXIST
        assert self.wb, "self.wb sould exist when we try to writeheader()."

    def close(self, save:bool=True):
        if save: self.save()
        if not self._saved: log.warning(f"Closing file {self.filename} without saving first.")
        if self.wb: self.wb = None
        self._opened = False

class rowIterator:
    def __init__(self, table):
        self.table = table
        self.maxlen = 50
    def __iter__(self):
        self.nrow = 1 # Start from line 2
        return self
    def __next__(self):
        if self.nrow <= self.maxlen:
            rv = self.table.get_row(self.nrow)
            self.nrow += 1
            return rv.get_values()        
        else:
          raise StopIteration

class ReadODS(ReadOnlySpreadsheet, ODSFile):
    # ------- ABSTRACT METHOD IMPLEMENTATIONS -----------
    def __init__(self:Self, filename: str|Path): 
        super().__init__(filename)
        self.wb = None
        self.table = None
        if self.fp.suffix.lower() not in SUFFIXES: 
            raise sheetError(f"{FORMAT_NAME} input file name must end in {SUFFIXES[0]}. {self.fp} fails")

    def open(self:Self) -> None:
        self._opened = False
        if not self.exists:
            raise sheetError(f"Coulf not access file {self.fp}.")
        self.open_ods()
        self._opened = True
        self.iter = iter(rowIterator(self.table))        

    def next(self:Self) :
        return self.iter
        
    def close(self:Self) -> None:
        if self.wb: self.wb = None
        self._opened = False
        
    @property
    def fields(self:Self)-> tuple[str,...]: return tuple(self._fields)

    @fields.setter
    def fields(self:Self, val) -> None:
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
    for d in t.next():
        print(tuple(d))
    t.close()
