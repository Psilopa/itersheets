# content of test_sysexit.py
from itersheets import itersheets
from pathlib import Path
import pytest

testfile = "test.csv"
class TestPropsClass:
    t = itersheets.ReaderCSV( testfile )
    def test_isopen(self):
        assert self.t.is_open == False
        self.t.open()
        assert self.t.is_open == True
    def test_filename(self):
        assert self.t.filename == Path(testfile)
    def test_exists(self):
        assert self.t.exists == True
    def test_fields(self):
        # We should write the test file to be sure
        assert self.t.fields == ("col_foo", "col_bar", "col_baz")
    def test_has_columns(self):
        assert self.t.has_columns(["col_foo","col_bar"]) == True
    def test_lacks_columns(self):
        assert self.t.lacks_columns(["col_baz","col_missing"]) == \
                    set(("col_missing",))
    def test_open(self):
        assert self.t.is_open == True
        self.t.close()
        assert self.t.is_open == False

