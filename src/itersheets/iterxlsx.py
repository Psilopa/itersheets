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

class XLSXFile():    
    def __init__(self:Self):
        self._reset()

    def _reset(self:Self):
        self.wb = None
        self.ws = None # ODS Table  equivalent
        
    def format_read_header(self:Self, filepath:Path) -> tuple:
        wb = load_workbook(filepath)
        assert wb, "wb should exist when we try to read a row in format_read_header()."
        ws = wb.active
        assert ws, "ws should exist when we try to read a row in format_read_header()."
        rv = tuple(x.value for x in ws[1:1])
        wb.close()
        return rv

    def format_open(self:Self, filepath: Path):
        if not filepath.exists(): # TODO: test for readable.
            raise sheetError(f"File {filepath} does not exist or is unavailable.")
        self.wb = load_workbook(filepath)        
        if not self.wb:
            raise sheetError(f"Could not open file {filepath}")
        self.ws = self.wb.active

    def format_create(self:Self):
        self.wb = Workbook()
        assert self.wb, "wb should exist when we try to read a row in format_create()."
        self.ws = self.wb.active

    def format_map_fields_to_columns(self:Self, fields:type_fields):
        return { k:v for k,v in zip(fields, range(len(fields))) } 
    
    def format_add_row(self:Self, values: Sequence):
        assert self.ws, "self.ws sould exist when we try to addrow()."
        self.ws.append( values )

    def format_get_row_values(self:Self, nrow:int) -> tuple:
        assert self.ws, "self.ws sould exist when we try to get row values."
        return tuple(x.value for x in ws[nrow+1:nrow+1])

    def format_set_row_values(self:Self, nrow:int, values:Sequence) -> None:
        max_col = len(values)
        assert self.ws, "self.ws sould exist when we try to writeheader()."
        for i in range(max_col):# 1-based indexing in Excel
            cell = self.ws.cell(row= nrow+1, column = i+1, value= values[i])
                
    def format_set_row_font(self:Self, nrow:int, font):
        assert self.ws, "self.ws sould exist when we try to writeheader()."
        for cell in self.ws[nrow+1:nrow+1]:# 1-based indexing in Excel
            cell.font = font

    def format_save(self:Self, filepath:Path) -> None:
        # TODO: ERROR HANDLING ON BAD PATH
        self.wb.save(filepath)

    def format_close(self:Self, filepath:Path=None) -> None:
        if self.wb and filepath:
            self.format_save(filepath)
        if self.wb:
            self.wb.close()
        self._reset()

class WriteXLSX(WriteOnlySpreadsheet, XLSXFile): 
    def __init__(self:Self, filename: str|Path): 
        super().__init__(filename)
        self.field2col = {} # Maps column name to column number
        if self.fp.suffix.lower() not in SUFFIXES: 
            raise sheetError(f"{FORMAT_NAME} output file name must end in {SUFFIXES[0]}. {self.fp} fails")

    # ------- ABSTRACT METHOD IMPLEMENTATIONS -----------
    def open(self:Self,
             fieldnames:Sequence[str] = tuple()
             ) -> None:
        self._opened = False        
        if not self.is_empty:
            self.format_open(self.fp)
             # If file exists, use first row as columns + any additional specified by useer
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
        self.format_add_row( outrow )
        self.save() # Saving after each insert if quite heavy
        
    def writeheader(self:Self) -> None:
        # TODO: ERROR STATE HANDLING
        self._saved = False
        assert self.ws, "self.ws sould exist when we try to writeheader()."
        self.format_set_row_values(0,self.fields)
        self.format_set_row_font(0, BOLDFONT)
        # Do not save just for a header, wait for a data row

    def close(self, save:bool=True):
        if save: self.save()
        if not self._saved: log.warning(f"Closing file {self.filename} without saving first.")
        self.wb.close()
        self._opened = False
        self.field2col = {} # Maps column name to column number

class ReadXLSX(ReadOnlySpreadsheet, XLSXFile):
    # ------- ABSTRACT METHOD IMPLEMENTATIONS -----------
    # Uses the openpyxl-provided iterator directly for row access
    def __init__(self:Self, filename: str|Path): 
        super().__init__(filename)
        self._reset_variables()
        if self.fp.suffix.lower() not in SUFFIXES: 
            raise sheetError(f"{FORMAT_NAME} inputput file name must end in {SUFFIXES[0]}. {self.fp} fails")

    def _reset_variables(self:Self):
        self._fields = tuple()

    def open(self:Self) -> None:
        self.format_open(self.fp)
        self.iter = self.ws.values
        self._fields = next(self.iter) # Grab 1st row and use it as field names

    def next(self:Self) :        
        return self.iter or tuple(tuple())# Return next row from the iterator

    def close(self:Self) -> None:
        self.format_close()
        
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
    for d in t.next():
        # Convert everything to strings and None to ""
        d = tuple(str(x) if x else "" for x in d)
        print(repr(d))
    t.close()
