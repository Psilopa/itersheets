# Itersheets
Iterative reading and writing of spreadsheets with a dictionary mapping approach: a file is assumed to have unique values in on the 1st line.

Targets multiple formats (so far, CSV and XLSX are supported). 

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
t.add_line_from_dict({"col_foo":"test1", "col_baz":"test2"}, on_new_key = itersheets.new_column_action.INSERT )

json = '{"col_foo":"jsonD",  "col_baz":"jsonE"}'
t.add_line_from_json(json, on_new_key = itersheets.new_column_action.INSERT )
t.save()
t.close()
```

## Accessing file and file content properties

```
print( "t.filename = ", t.filename )
print( "t.is_open = ", t.is_open )
print( "t.exists = ", t.exists )
print( "t.is_empty = ", t.is_empty )
print( "t.fields = ", t.fields )
print( "t.has_columns = ", t.has_columns(["col_foo","col_bar"]) ) # Returns True
print( "t.has_columns = ", t.has_columns(["col_baz","col_missing"]) ) # Returns False
print( "t.lacks_columns = ", t.lacks_columns(["col_baz","col_missing"]) ) # Returns a set of missing columns
print( "t.is_open = ", t.is_open )
```
