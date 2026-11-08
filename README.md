# Itersheets
Reading and writing of spreadsheets with a dictionary mapping approach. A file is assumed to have unique values in on the 1st line.
The available functionality is limited and most restriced to basic data reading and writing: the main purpose of this module is to allow easy cross-format support for these basic tasks.

Targets multiple formats. CSV XLSX, and ODS are currently supported. 

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
t.open(["col_foo", "col_bar"])
t.add_row_from_dict({"col_foo":"test1", "col_baz":"test2"}, on_new_key = itersheets.new_column_action.INSERT )

json = '{"col_foo":"jsonD",  "col_baz":"jsonE"}'
t.add_row_from_json(json, on_new_key = itersheets.new_column_action.INSERT )
t.save()
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
