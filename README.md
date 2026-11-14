# Itersheets
Reading and writing of basic tabular data line-line line with an uniform API for CVS, ODS and XLSX files.

As the name suggests, all row operations are done in one sweep row-by-row. No random access is allowed.

**Work in progress for now**! The current version assumes files have unique values in on the 1st line (which allows output as dictionaries). Future versions can hopefully release this condition or make it optional.

# TODO
- Allow users to specify first line of data
- ? Allow multiple header lines, if >1, join column headers as a tuple for dictionary access (header_row1, header_row2): data
- Support CSV in-place line addition

# Installation
For now, you have to build your own Python package if you want to use this as a simple import. 

- Download source code 
- ```pip install --upgrade setuptools[core]```
- In the directory where the pyconfig.toml file is, run ```python -m build```
- ```pip install dist/itersheets-0.4-py3-none-any.whl```  (or whatever is the name fo the .whl package file created by setuptools

# Basic usage

## Writing data

```
from itersheets import itersheets
t = itersheets.WriterXLSX( "test.xlsx" )
t.open("w", ["col_foo", "col_bar"])
t.add_row_from_dict({"col_foo":"test1", "col_baz":"test2"}, on_new_key = itersheets.new_column_action.INSERT )

json = '{"col_foo":"jsonD",  "col_baz":"jsonE"}'
t.add_row_from_json(json, on_new_key = itersheets.new_column_action.INSERT )
t.save()
t.close()
```

## Reading data

```
t = itersheets.Reader( "test.xlsx" ) # Generic reader guesses the type from the file name
t.open()
print(f"Header fields (1st row values):", t.fields)
for d in t.dictrows():
    print("row as dictionary = ", d)
t.close()
```

## Accessing file and file content properties

```
print( "t.filename = ", t.filename ) # Path
print( "t.is_open = ", t.is_open ) # Boolean 
print( "t.exists = ", t.exists ) # Boolean
print( "t.is_empty = ", t.is_empty ) # Boolean
print( "t.fields = ", t.fields ) # List of field names (=unique values on the 1st row)
print( "t.has_columns = ", t.has_columns(["col_foo","col_bar"]) ) # Boolean, True with the file written above
print( "t.has_columns = ", t.has_columns(["col_baz","col_missing"]) ) # Boolean, False with the file written above
print( "t.lacks_columns = ", t.lacks_columns(["col_baz","col_missing"]) ) # Returns a set of missing columns

```
