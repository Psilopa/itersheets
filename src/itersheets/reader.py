# Reader classes for ODS and XLSX (CSV is quite different and in its own file)
# STDlib
from pathlib import Path

# My own modules
from .errors import sheetError
from .base import ReadOnlySpreadsheet
from .jkxlsx import XLSXFile
from .jkods import ODSFile

class ReaderBase(ReadOnlySpreadsheet):
    def __init__(self:Self, filename: str|Path): 
        super().__init__(filename)
        self._reset_variables()

    def _reset_variables(self:Self):
        self._opened = False
        self.fields = tuple()
        self.file = None

    # ------- ABSTRACT METHOD IMPLEMENTATIONS -----------
    def open(self:Self) -> None:
        self._opened = False
        if not self.exists:
            raise sheetError(f"Coulf not access file {self.fp}.")
        self.format_open(self.fp)
        self.iter = self.value_iterator() 
        self.fields = next(self.iter) # Grab 1st row and use it as field names
        self._opened = True

    def __next__(self:Self):
        return next(self.iter) # Return next row from the iterator or an empty tuple

    def close(self:Self) -> None:
        self.format_close()
        self._reset_variables()

#------------------ EXCEL SPECIFIC STUFF -------------
class ReaderXLSX(ReaderBase, XLSXFile):
    def __init__(self:Self, filename: str|Path): 
        ReaderBase.__init__(self,filename)
        XLSXFile.__init__(self)

class ReaderODS(ReaderBase, ODSFile):
    def __init__(self:Self, filename: str|Path): 
        ReaderBase.__init__(self,filename)
        XLSXFile.__init__(self)
    def open(self:Self) -> None:
        ReaderBase.open(self)
        # HACK: for some reaso, the iter in ODS gets started 2,
        # but in XLSX, only once?
        # Resulting in an off-by-1 difference
        # So we set fields manually here and jkods.rowIterator starts from row 2
        self.fields = self.format_get_row_values(0)       
