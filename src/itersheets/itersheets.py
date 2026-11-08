# Expose main classes
from writer import WriterXLSX, WriterODS
from reader import ReaderXLSX, ReaderODS
from jkcsv import WriterCSV, ReaderCSV
from base import new_column_action


#------------------ MAIN FOR TESTING -------------
def _test(tclasses, testfile, tformat, on_new_key):
    writeclass, readclass = tclasses
    print(f"---------------TESTING {tformat} .............:")
    # XLSX Writing
    print(f"TEST Writing {tformat}.")
    t = writeclass( testfile )
    t.open(["col_foo", "col_bar"])
    t.add_line_from_dict({"col_foo":"test1", "col_bar":"test2"}, on_new_key=on_new_key)
    t.add_line_from_dict({"col_foo":"testA", "col_bar":"testB"}, on_new_key=on_new_key)
    t.add_line_from_dict({"col_foo":"testC", "col_bar":"testC",
                          "col_baz":"testC"}, on_new_key=on_new_key)
    t.add_line_from_dict({"col_foo":"testD",  "col_baz":"testD"}, on_new_key=on_new_key)
    t.save()
    t.close()
    # XLSX Reading
    testformat = "XLS"
    print(f"TEST READING {tformat}.")
    t = readclass( testfile )
    t.open()
    print(f"Header fields in the {tformat} file:", t.fields)
    for d in t.next():
        # Convert everything to strings and None to ""
        d = tuple(str(x) if x else "" for x in d)
        print(repr(d))
    t.close()
    
if __name__ == '__main__':
    _test((WriterXLSX, ReaderXLSX), "test.xlsx", "Excel", new_column_action.INSERT)
    _test((WriterODS, ReaderODS), "test.ods", "ODS", new_column_action.INSERT)
    # For now, CSV can only swap to a new file on new field insert
    _test((WriterCSV, ReaderCSV), "test.csv", "CSV", new_column_action.NEWFILE)
