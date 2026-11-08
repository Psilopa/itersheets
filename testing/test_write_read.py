from itersheets import itersheets

def _test_write_read(tclasses, testfile, tformat, on_new_key):
    writeclass, readclass = tclasses
    print(f"---------------TESTING {tformat} .............:")
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
    # Reading
    testformat = "XLS"
    print(f"TEST READING {tformat}.")
    t = readclass( testfile )
    t.open()
    print(f"Header fields in the {tformat} file:", t.fields)
    for d in t.next():
        # Convert everything to strings and None to ""
        d = tuple(str(x) if x else "" for x in d)
        print(repr(d))
    # Properties
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
    
if __name__ == '__main__':
    # Define actions in case a new column after the first row to data was written.
    insert = itersheets.new_column_action.INSERT
    newfile = itersheets.new_column_action.NEWFILE
    _test_write_read((itersheets.WriterXLSX, itersheets.ReaderXLSX), "test.xlsx", "Excel", insert)
    _test_write_read((itersheets.WriterODS, itersheets.ReaderODS), "test.ods", "ODS", insert)
    # For now, CSV can only swap to a new file on new field insert
    _test_write_read((itersheets.WriterCSV, itersheets.ReaderCSV), "test.csv", "CSV", newfile)
