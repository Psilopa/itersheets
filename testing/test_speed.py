from itersheets import itersheets
import timeit
# Read 


def test_read():
        global fp
        global dicts
        t = itersheets.Reader(fp)
        t.open()
        dicts = []
        l = 0
        for d in t.dictrows():
#                print(f"Data line :", d)
#                print(".")
                dicts.append(d)
                l += 1
        print(f"Header fields (1st row values):", t.fields)
        print(f"Read {l} lines.")
        t.close()

def test_write():
        global fp
        global dicts
        global newcol 
        t = itersheets.Writer(fp)
        t.open("a")
        for d in dicts:
                t.add_row_from_dict(d, on_new_key =  newrow[i])
        t.save()
        t.close()

# Test reading
dicts = []
formats = (".csv", ".xlsx", ".ods")
newrow = (itersheets.new_column_action.NEWFILE,
          itersheets.new_column_action.INSERT,
          itersheets.new_column_action.INSERT)
fbase = "needscompletion_300k_v1_T"

do = (True, True, True )  # False
for i in range(3):	
	if do[i]:
		fformat = formats[i]
		fp = fbase + fformat
		timer = timeit.Timer(test_read)
		elapsed = timer.timeit(1)
		print(f'Read time for {fformat}: {elapsed:.6f} seconds')
#print(dicts)
# Test writing
fbase = "out"
do = (True, True, True )  # False
for i in range(3):	
	if do[i]:
		fformat = formats[i]
		fbase = "out"
		fp = fbase + fformat
		print("Writing to", fp)
		timer = timeit.Timer(test_write)
		elapsed = timer.timeit(1)
		print(f'Write time for {fformat}: {elapsed:.6f} seconds')

