# This is a sample Python script.
import pandas as pd
import matplotlib.pyplot as plt
RosiePath = 'U:/FAC/Sales Support_PL/9_FAC SUPPORT/5_FILES/ROSIE/Analysis/validationStatus5.xlsx'
print("Rosie Data Loading:")
RosieData = pd.read_excel(RosiePath,
                        usecols=['quote number', 'pos', 'sapExText', 'decision', 'OPCcaseCreated', 'validation status'],
                        squeeze=True,
                        sheet_name='Sheet1'
                          )
OPCNotCreated = RosieData['OPCcaseCreated'] == 'No'
isApproved = RosieData['decision'] == 'Yes'

containsElement = RosieData['sapExText'].str.contains('Element')

isV200 = RosieData['sapExText'].str.contains('Id:V200') | \
         RosieData['sapExText'].str.contains('Id:V200E') | \
         RosieData['sapExText'].str.contains('Id:V200i') | \
         RosieData['sapExText'].str.contains('Id:VELFAC_EDG')

# isV200 =        RosieData['sapExText'].str.contains('Id:V200')
# isV200E =       RosieData['sapExText'].str.contains('Id:V200E')
# isV200i =       RosieData['sapExText'].str.contains('Id:V200i')
# isVELFAC_EDG =  RosieData['sapExText'].str.contains('Id:VELFAC_EDG')

isGP21 = RosieData['sapExText'].str.contains('Id:RIBO_ALU_2') | \
        RosieData['sapExText'].str.contains('Id:RIBO_WOOD') | \
         RosieData['sapExText'].str.contains('Id:AURA2015') | \
         RosieData['sapExText'].str.contains('Id:AURAPLUS20') | \
         RosieData['sapExText'].str.contains('Id:CLASSIC_AL') | \
         RosieData['sapExText'].str.contains('Id:CLASSIC_WO') | \
         RosieData['sapExText'].str.contains('Id:FORMA2015') | \
         RosieData['sapExText'].str.contains('Id:FORMAPLUS2')
lSystems = [isV200, isGP21]
sSystemsNames = ['V200', 'GP21']
isKrone = RosieData['sapExText'].str.contains('Id:VELFAC_IN') | RosieData['sapExText'].str.contains('Id:AURAPLUS_I')

isVelfacIn = RosieData['sapExText'].str.contains('Id:VELFAC_IN')
isAURAPLUS_I = RosieData['sapExText'].str.contains('Id:AURAPLUS_I')

isStatusEqualFive = RosieData["validation status"] == 5
isTestQuote = RosieData['quote number'] == 20139832
isTITUD = RosieData['sapExText'].str.contains('TITUD')

lOpeningFunctions = ['SHO', 'SGO', 'SHRO', 'FL', 'FC', 'SCD:SR','SCD:SL','SCD:R','SCD:L', 'THRO', 'THROX', 'THO', 'TGO', 'FCC', 'BHI', 'BHO', 'TUD',
                     'TITUD', 'TITU', 'CDO', 'PDO', 'PDI', 'EDO', 'EDI', 'Element']
#
# # print(RosieData[isVelfacIn & isStatusEqualFive].head())
# result = RosieData[isTITUD]
# float(s[s.find('W:') + 2:s.find('H:', s.find('W:') + 2)].replace('.','').replace(',','.'))
# RosieData['sapExText']

# kroneTITUD = RosieData[(isVelfacIn | isAURAPLUS_I) & isTITUD]
# kroneTITUD[['iD', 'opening', 'Width', 'Height', 'SashWeight']] = kroneTITUD['sapExText'].str.split(' ', expand=True)
#
#
# kroneTITUD[['Width1', 'Width2']] = kroneTITUD['Width'].str.split(':', expand=True)
# kroneTITUD['Width'] = kroneTITUD['Width2'].str.replace('.','').str.replace(',','.').astype(float)
#
# kroneTITUD[['Height1', 'Height2']] = kroneTITUD['Height'].str.split(':', expand=True)
# kroneTITUD['Height'] = kroneTITUD['Height2'].str.replace('.','').str.replace(',','.').astype(float)


# plt.scatter(kroneTITUD['Height'], kroneTITUD['Width'])
# plt.show()

# jakaś nieudana próba, ale wynik jest interesujący
# kroneTITUTD['Width'] = \
#     float(str(kroneTITUTD['sapExText'])[str(kroneTITUTD['sapExText']).find('W:') + 2:
#                                  str(kroneTITUTD['sapExText']).find('H:',
#                                  str(kroneTITUTD['sapExText']).find('W:') + 2)]
#                                  .replace('.','').replace(',','.'))
# print(kroneTITUTD)
# print(result)
outFile = 'C:/Users/bzm.fac/Dropbox/Python/TableAnalysis/outFile2.xlsx'
# for system in lSystems:
#     print(RosieData[system].head())
#
# for opening in lOpeningFunctions:
#     print(RosieData[RosieData['sapExText'].str.contains(opening)].head())
# writer = pd.ExcelWriter(outFile, engine='xlsxwriter')
iSystemCount = 0
for system in lSystems:
    iOpeningCount = 0
    for openning in lOpeningFunctions:
        fOpening = RosieData['sapExText'].str.contains(openning)
        if not RosieData[OPCNotCreated & system & fOpening].empty:
            print('System: '+ sSystemsNames[iSystemCount]+
                  ' Opening: '+ lOpeningFunctions[iOpeningCount] +
                  ' Items in status 5: ' + str(len(RosieData[isApproved & OPCNotCreated & system & fOpening])))
            sheet_name = (sSystemsNames[iSystemCount]+lOpeningFunctions[iOpeningCount].replace(':',''))

            with pd.ExcelWriter(outFile, mode='a') as writer:
                RosieData[isApproved & OPCNotCreated & system & fOpening].to_excel(writer, sheet_name=sheet_name)

        iOpeningCount += 1
    iSystemCount +=1

# iSystemCount = 0
# for system in lSystems:
#     fSystem = RosieData[system]
#     iOpeningCount = 0
#     for opening in lOpeningFunctions:
#         fOpening = RosieData['sapExText'].str.contains(str(opening))]
#         print(fOpening)
#         state = RosieData[fSystem & fOpening].empty
#         print(state)
#         if len(RosieData[fSystem & fOpening]) > 0:
#             with pd.ExcelWriter(outFile) as writer:
#                 RosieData[fSystem & fOpening].to_excel(writer,
#                                                 sheet_name=sSystemsNames(iSystemCount)+lOpeningFunctions(iOpeningCount))
#         iOpeningCount += 1
#     iSystemCount +=1



# with pd.ExcelWriter(outFile) as writer:
#
#     dlaPJJ.to_excel(writer, sheet_name="Sheet1")
#
#     dlaPJJ.to_excel(writer, sheet_name="Sheet2")

# rysowanie linii:
#
# # draw vertical line from (70,100) to (70, 250)
# plt.annotate("",
#               xy=(700, 1000), xycoords='data',
#               xytext=(700, 2500), textcoords='data',
#               arrowprops=dict(arrowstyle="-",
#                               connectionstyle="arc3,rad=0."),
#               )

# draw diagonal line from (70, 90) to (90, 200)
# plt.annotate("",
#               xy=(700, 900), xycoords='data',
#               xytext=(900, 2000), textcoords='data',
#               arrowprops=dict(arrowstyle="-",
#                               connectionstyle="arc3,rad=0."),
#               )
# #
# plt.show()

# s = 'Id:VELFAC_IN fl:1TITUD:L W:1.053,0 H:2.043,0 Sh.Wgt:75,29'
# # start = s.find('W:') + 2
# # end = s.find('H:', start)
# # result = s[s.find('W:') + 2:s.find('H:', s.find('W:') + 2)]
# result2 = float(s[s.find('W:') + 2:s.find('H:', s.find('W:') + 2)].replace('.','').replace(',','.'))
# print(s)
# print('Typ:', type(result2), 'Wynik:', result2)
print("End Program!")


