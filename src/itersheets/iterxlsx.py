# STDlib
import logging, csv, re
from pathlib import Path

# Non-stdlib
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

# My own modules
from errors import sheetError
from base import new_column_action, WriteOnlySpreadsheet, ReadOnlySpreadsheet
#from .mytyping import * # For type checking tools

BOLDFONT = Font(bold=True)
SUFFIXES = [".xlsx", "xlsx"]
FORMAT_NAME = "Excel"

log = logging.getLogger() # Set from main script if fancy logging is required.

class WriteXLSX(WriteOnlySpreadsheet): 
    def __init__(self:Self, filename: str|Path): 
        super().__init__(filename)
        self.field2col = {} # Maps column name to column number
        if self.fp.suffix.lower() not in SUFFIXES: 
            raise sheetError(f"{FORMAT_NAME} output file name must end in {SUFFIXES[0]}. {self.fp} fails")
    # ------- ABSTRACT METHOD IMPLEMENTATIONS -----------
    def _read_header(self:Self) -> tuple:
        wb = load_workbook(self.fp)
        ws = wb.active
        assert ws, "ws should exist when we try to read a row in _read_header()."
        iterator = ws.iter_rows(min_row=1, max_row=1, values_only=True) # 1-based indexing, Excel-style
        rv = next(iterator)
        wb.close()
        return rv

    def open(self:Self,
             fieldnames:Sequence[str] = tuple()
             ) -> None:
        # If file exists, use first row as columns
        if not self.is_empty:
            # Get column data from existing file
            self.fields = self._read_header() + tuple(fieldnames)
            self.wb = load_workbook(self.fp)
            self.ws = self.wb.active
        else:
            self.fields = fieldnames
            self.wb = Workbook()
            self.ws = self.wb.active
#            self.ws.title = ""
        # Map field name to column number (0-based indexing as we do not use this numbering directly in Excel)
        self.field2col = { k:v for k,v in zip(self.fields, range(len(self.fields))) } 
        print(self.field2col)
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
        assert self.ws, "self.ws sould exist when we try to addrow()."
        self.ws.append( outrow )
        self.save() # Saving after each insert if quite heavy
        
    def writeheader(self:Self) -> None:
        # TODO: ERROR STATE HANDLING
        # TODO: THIS SHOULD ALWAYS ACCESS LINE 1
        max_col = len(self.fields)
        n = 0
        self._saved = False
        assert self.ws, "self.ws sould exist when we try to writeheader()."
        for col in self.ws.iter_cols(min_row=1, max_col=max_col, max_row=1): # 1-based indexing in Excel
            for cell in col:
                cell.value = self.fields[n]
                cell.font = BOLDFONT
                n += 1

    def close(self, save:bool=True):
        if save: self.save()
        if not self._saved: log.warning(f"Closing file {self.filename} without saving first.")
        self.wb.close()
        self._opened = False

class ReadXLSX(ReadOnlySpreadsheet):
    # ------- ABSTRACT METHOD IMPLEMENTATIONS -----------
    def __init__(self:Self, filename: str|Path): 
        super().__init__(filename)
        self.wb = None
        self.ws = None
        if self.fp.suffix.lower() not in SUFFIXES: 
            raise sheetError(f"{FORMAT_NAME} inputput file name must end in {SUFFIXES[0]}. {self.fp} fails")

    def open(self:Self) -> None:
        self._opened = False
        if not self.exists:
            raise sheetError(f"Coulf not access file {self.fp}.")
        self.wb = load_workbook(self.fp)
        self.ws = self.wb.active
        if not self.ws:
            raise sheetError("Could not accewss file {self.filename}")
        self.iter = self.ws.values
        self._fields = next(self.iter) # Grab 1st row and use it as field names
        self._opened = True

    def next(self:Self) :        
        return self.iter or tuple(tuple())# Return next row from the iterator

    def close(self:Self) -> None:
        if self.wb: self.wb.close()
        self._opened = False
        
    @property
    def fields(self:Self)-> tuple[str,...]: return tuple(self._fields)

if __name__ == '__main__':
    testfile = "test.xlsx"
    t = WriteXLSX( testfile )
    t.open(["col_foo", "col_bar"])
    t.add_line_from_dict({"col_foo":"test1", "col_bar":"test2"}, on_new_key=new_column_action.INSERT)
    t.add_line_from_dict({"col_foo":"testA", "col_bar":"testB"}, on_new_key=new_column_action.INSERT)
    t.add_line_from_dict({"col_foo":"testB", "col_bar":"testB", "col_baz":"testB"}, on_new_key=new_column_action.INSERT)
    t.add_line_from_dict({"col_foo":"testB",  "col_baz":"testB"}, on_new_key=new_column_action.INSERT)
    t.save()
    t.close()
    t = ReadXLSX( testfile )
    t.open()
    print("HEADER LINE", t.fields)
    for d in t.next():
        # Convert everything to strings and None to ""
        d = tuple(str(x) if x else "" for x in d)
        print(repr(d))
    t.close()
