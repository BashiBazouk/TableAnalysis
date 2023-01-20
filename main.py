# imports
import pandas as pd

# hardcodes

input_file = 'validationStatus5.xlsx'
output_file = 'outFile2.xlsx'

print("Loading data from excel now")

# creating pandas dataframe - load from input file
DataFrame = pd.read_excel(input_file,
                        usecols=['quote number', 'pos', 'foo_text', 'decision', 'request_created', 'validation status'],
                        squeeze=True,
                        sheet_name='Sheet1'
                          )
                          
# example of filters:
# bool filters
isRequest_created = DataFrame['request_created'] == 'No'
isApproved = DataFrame['decision'] == 'Yes'

# filter with single element/hit
containsElement = DataFrame['foo_text'].str.contains('Element')

#filters with multiple elements/hits
isFoo_1 = DataFrame['foo_text'].str.contains('foo_1') | \
         DataFrame['foo_text'].str.contains('foo_1E') | \
         DataFrame['foo_text'].str.contains('foo_1i') | \
         DataFrame['foo_text'].str.contains('foo_1_EDG')

isFoo_2 = DataFrame['foo_text'].str.contains('foo_2') | \
         DataFrame['foo_text'].str.contains('foo_2E') | \
         DataFrame['foo_text'].str.contains('foo_2i') | \
         DataFrame['foo_text'].str.contains('foo_2_EDG')

# list of systems to iterate to:
lSystems = [isFoo_1, isFoo_2]

sSystemsNames = ['Foo_1', 'Foo_2']

# list of opening functions: 
# todo move lists to external file, for clarification and ease of editing
lOpeningFunctions = ['bar_1', 'bar_2', 'bar_3', 'bar_4']

# iterate over system, over opening function
iSystemCount = 0
for system in lSystems:
    iOpeningCount = 0
    for openning in lOpeningFunctions:
        fOpening = DataFrame['foo_text'].str.contains(openning)
        if not DataFrame[isRequest_created & system & fOpening].empty:

            # some terminal output for quick check how script works
            print('System: '+ sSystemsNames[iSystemCount]+
                  ' Opening: '+ lOpeningFunctions[iOpeningCount] +
                  ' Items in status 5: ' + str(len(DataFrame[isApproved & isRequest_created & system & fOpening])))

            # determine sheet name - system + opening = foo_1bar_1
            sheet_name = (sSystemsNames[iSystemCount]+lOpeningFunctions[iOpeningCount]

            # open file and write filtered data frame to new excel sheet
            # I think that with statement should be moved before loops #todo
            with pd.ExcelWriter(output_file, mode='a') as writer:
                DataFrame[isApproved & isRequest_created & system & fOpening].to_excel(writer, sheet_name=sheet_name)

        iOpeningCount += 1
    iSystemCount +=1

print("End Scrip!")
print("Have a nice day!")