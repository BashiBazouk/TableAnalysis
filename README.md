# Move data from original sheet to subsheets - excel but with Python Pandas

## Purpose
I've had a repetetive analysis task - create sheets with data from original report from another software. Filter out interesting data, move it to another tab, write down some results, rinse and reapeat. 
Why not do it in different way? Why not learn something new?
I've did few scripts with Python Pandas library, it seems to be ideal opportunity to do something new.

## Description, How it works
Instead of making combinations of filters in excel and then copying filtered data we can make algorithm of this. Simplified structure of data is:<br>
- system1
	- opening1
	- opening2
	- ...
- system2
	- opening1
	- opening2
	- ...
- ...

as a result we should have excel with sheets: `system1_opening1`, `system1_opening2`, ...<br>
this is how to do it with [Pandas](https://pandas.pydata.org/)<br><br>


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

Now just wrap it in two loops (as my data was only two levels deep):
for system in systems
	for opening in openings
		DataFrame[



1. Project's Title
2. Project Description

- What your application does,
- Why you used the technologies you used,
- Some of the challenges you faced and features you hope to implement in the future.

