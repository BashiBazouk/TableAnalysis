# Copy data from original sheet to subsheets based on data in columns - excel filtering but with Python Pandas

## Purpose
I've had a repetetive analysis task - create sheets with data from original report from another software. Unfortunately data is not easy to filter, as interestind data is in one long string of data. At start one column must be splitted in few more columns. Filter out interesting data, move it to another tab, write down some results, rinse and reapeat. 
Why not do it in different way? Why not learn something new?
I've did few scripts with Python Pandas library, it seems to be ideal opportunity to do something new.

## Description, How it works
Instead of making combinations of filters in excel and then copying filtered data we can make algorithm of this.<br>
 Simplified structure of data is:<br>
`data`<br>
`  |`<br>
`  +--->system1`<br>
`  |      |`<br>
`  |      +--->opening1`<br>
`  |      |`<br>
`  |      +--->opening2`<br>
`  |      |`<br>
`  |      +--->...`<br>
`  |`<br>
`  +--->system2`<br>
`  |      |`<br>
`  |      +--->opening1`<br>
`  |      |`<br>
`  |      +--->opening2`<br>
`  |      |`<br>
`  |      +--->...`<br>
`  |`<br>
`  +--->...`<br>
so, go in original sheet1, filter by `system1`, look for `opening1`, export to new tab, calculate how many rows, filter by `system1`, look for `opening2` ...<br>
as a result we should have excel with sheets: `system1opening1`, `system1opening2`, ...<br>

## How to script it:

1. First of all we need to load data from excel to Pandas data frame object like this:<br><br>
`table = pandas.read_excel(file, usecols=['col1','col2'], sheetname='Sheet1')`<br>
`usecols=` let you choose which columns to load into table, usefull when your excel has a lot of columns which you will not use<br>
`Sheetname=` the same, but for sheets
more info in<br>
[official Pandas documentation](https://pandas.pydata.org/docs/reference/api/pandas.read_excel.html)<br>

2. Seting filters<br>
just like in excel filter can be set to look for strings. In Pandas:
`isOpening1 = DataFrame['ColumnWithOpenings'].str.contains('opening1')`<br>
ypu can combine filters wit python operators:
`isWindowSystem = DataFrame['ColumnWithWindowSystem'].str.contains('system1') |`<br>
`		  DataFrame['ColumnWithWindowSystem'].str.contains('system2')`<br>
when you apply filters to data frame you will get only data (another data frame) with only desired data.<br>
such data frame can be saved as a new sheet.
3. Loops<br>
Now just wrap it in two loops (as my data was only two levels deep):<br>
`for system in systems`<br>
	`for opening in openings`<br>
		`DataFrame[filters].to_excel()`<br>
<br>
Please check<br> [Pandas saving data frame to excel](https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.to_excel.html) <br>

as a result you'll get new sheet only with rows which fit the filters.<br>
at this point by changing filters you can filter out intersting data in seconds.<br>

## todo
[ ] remove hardcoding
[ ] check if "with" statement should be moved before loops. It's working, but it's quite slow at this moment
[ ] move filters and list to separate file, not to edit main script all the time
[ ] add "try catch" for now I need to take care that all data is in place before running script
