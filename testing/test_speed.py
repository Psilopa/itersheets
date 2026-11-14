from itersheets import itersheets
import timeit
# Read 


def test_read():
	global fp
	global dicts
	t = itersheets.Reader( fp)
	t.open()
	dicts = []
	print(f"Header fields (1st row values):", t.fields)
	for d in t.dictrows():
		dicts.append(d)
	t.close()


def test_write():
	global fp
	global dicts
	t = itersheets.Writer( fp)
	t.open()
	for d in dicts:
		t.add_row_from_dict(d, on_new_key = itersheets.new_column_action.INSERT )
	t.save()
	t.close()

dicts = []
formats = (".csv", ".xlsx", ".ods")
fbase = "needscompletion_300k_v1_T"
do = (True, True, False )  # False

for i in range(3):	
	if do[i]:
		fformat = formats[i]
		fp = fbase + fformat
		timer = timeit.Timer(test_read)
		elapsed = timer.timeit(1)
		print(f'Read time for {fformat}: {elapsed:.6f} seconds')

fbase = "out"
do = (False, True, False )  # False
for i in range(3):	
	if do[i]:
		fformat = formats[i]
		fbase = "out"
		fp = fbase + fformat
		timer = timeit.Timer(test_write)
		elapsed = timer.timeit(1)
		print(f'Write time for {fformat}: {elapsed:.6f} seconds')

