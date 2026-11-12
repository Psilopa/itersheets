from itersheets import itersheets
from pathlib import Path

def test_write(writeclass, testfile, tformat, on_new_key):
    print(f"---------------TESTING  .............:")
    # Writing
    print(f"TEST Writing {tformat}.")
    t = writeclass( testfile )
    t.open(["col_foo", "col_bar"])
    t.add_row_from_dict({"col_foo":"test1", "col_bar":"test2"}, on_new_key=on_new_key)
    t.add_row_from_dict({"col_foo":"testA", "col_bar":"testB"}, on_new_key=on_new_key)
    t.add_row_from_dict({"col_foo":"testC", "col_bar":"testC",
                          "col_baz":"testC"}, on_new_key=on_new_key)
    json = '{"col_foo":"jsonD",  "col_baz":"jsonE"}'
    t.add_row_from_json(json, on_new_key=on_new_key)
    t.save()
    t.close()
    print("END OF TEST\n")
    
def test_read(readclass, testfile):
    print(f"---------------TESTING  .............:")
    # Reading
    tformat = Path(testfile).suffix.strip(".")
    print(f"Test readlists lists from {tformat}.")
    t = readclass( testfile )
    t.open()
    print(f"Header fields in the {tformat} file:", t.fields)
    for d in t:
        # Convert everything to strings and None to ""
        print("row = ", d)
        d = tuple(str(x) if x else "" for x in d)
    t.close()
    print("END OF TEST\n")

def test_read_dict(readclass, testfile):
    print(f"---------------TESTING  .............:")
    # Reading
    tformat = Path(testfile).suffix.strip(".")
    print(f"TEST Reading dictionaries from {tformat}.")
    t = readclass( testfile )
    t.open()
    print(f"Header fields in the {tformat} file:", t.fields)
    for d in t.dictrows():
        # Convert everything to strings and None to ""
        print("row = ", d)
        d = tuple(str(x) if x else "" for x in d)
    t.close()
    print("END OF TEST\n")

def test_properties(readclass, testfile):
    print(f"---------------TESTING  .............:")
    print("Properties in  {tformat}")
    # Reading
    tformat = Path(testfile).suffix.strip(".")
    t = readclass( testfile )
    t.open()
    print("sheet.filename = ", t.filename)
    print("sheet.is_open = ", t.is_open)
    print("sheet.exists = ", t.exists)
    print("sheet.is_empty = ", t.is_empty)
    print("sheet.fields = ", t.fields)
    print("sheet.has_columns = ", t.has_columns(["col_foo","col_bar"]) )
    print("sheet.has_columns = ", t.has_columns(["col_baz","col_missing"]) )
    print("sheet.lacks_columns = ", t.lacks_columns(["col_baz","col_missing"]) )
    print("sheet.is_open = ", t.is_open)
    t.close()
    print("END OF TEST\n")


if __name__ == '__main__':
    # Define actions in case a new column after the first row to data was written.
    insert = itersheets.new_column_action.INSERT
    newfile = itersheets.new_column_action.NEWFILE
#    test_write(itersheets.WriterXLSX, "test.xlsx", "Excel", insert)
#    test_write(itersheets.WriterODS, "test.ods", "ODS", insert)
#    test_write(itersheets.WriterCSV, "test.csv", "CSV", newfile)
#    test_read( itersheets.ReaderCSV, "test.csv")
#    test_read_dict( itersheets.ReaderCSV, "test.csv")
#    test_properties( itersheets.ReaderCSV, "test.csv")
#    test_read( itersheets.ReaderXLSX, "test.xlsx")
#    test_read_dict( itersheets.ReaderXLSX, "test.xlsx")
    test_properties( itersheets.ReaderXLSX, "test.xlsx")
    test_read( itersheets.ReaderXLSX, "test.xlsx")
    test_read_dict( itersheets.ReaderXLSX, "test.xlsx")
