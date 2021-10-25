'''
    Skellig
    DeltaV fhx Data Extraction Tool
    Rev: 0.0
    
    This tool extracts data from fhx file and builds a spreadsheet containing the following data:
     - Phase class step, action, and transition data
     - Phase class variable data
     - Equipment module class step, action, and transition data
     - Equipment module class variable data
     - Named set data
'''

#Importing python libraries
import pandas as pd
import fhxutilities as util
from time import strftime

#Import config files
import fhxconstants as const

#Declare constants
if True:
    CLASSES, NAMED_SETS, FB_DEF, FBS = const.CLASSES, const.NAMED_SETS, const.FB_DEF, const.FBS

#Prompt for module file name
fileName = input('\nEnter name of the DeltaV Export (.fhx) file without file extension.\n') + '.fhx'

#Save start time
startTime = {'Hours':int(strftime("%H")),
             'Minutes':int(strftime("%M")),
             'Seconds':int(strftime("%S"))}

#Build list of strings from file
fhxLines = util.BuildLinesFromFhx(fileName)

numOfLines = len(fhxLines)

#Save sections of fhx as paragraph objects
[classes, namedSets, fbDefinitions, fbInstances] = util.SaveParagraphs(fhxLines, [CLASSES, NAMED_SETS, FB_DEF, FBS])

#Build named set dictionary and dataframe
namedSetMap = util.BuildNamedSetData(fhxLines, namedSets)

namedSetData = util.BuildNamedSetDataframe(namedSetMap)

#Builds EM and phase variable dataframe
emVarData, phaseVarData = util.BuildVariableDataframe(fhxLines, classes)

#Builds class info and command composite mapping dataframe
classCompMap = util.BuildClassCompData(fhxLines, classes, namedSetMap, fbInstances)

#Generates SFC dataframes and object order
emSFCData = util.BuildSFCDataframe(fhxLines, fbDefinitions, classCompMap, const.EM_CLASS)
phaseSFCData = util.BuildSFCDataframe(fhxLines, fbDefinitions, classCompMap, const.PHASE_CLASS)

#Writes dataframes to new spreadsheet
print('Building spreadsheet...')

fileName = fileName[:-4] + '_Steps&Transitions_' + strftime("%Y%m%d-%H%M%S") + '.xlsx'

with pd.ExcelWriter(fileName) as spreadsheet:
    emSFCData.to_excel(spreadsheet, sheet_name='EM SFC', freeze_panes=(1,0))
    phaseSFCData.to_excel(spreadsheet, sheet_name='Phase SFC', freeze_panes=(1,0))
    emVarData.to_excel(spreadsheet, sheet_name='EM Variables', freeze_panes=(1,0))
    phaseVarData.to_excel(spreadsheet, sheet_name='Phase Variables', freeze_panes=(1,0))
    namedSetData.to_excel(spreadsheet, sheet_name='Named Sets', freeze_panes=(1,0))

#Saves time duration
endTime = {'Hours':int(strftime("%H")),
           'Minutes':int(strftime("%M")),
           'Seconds':int(strftime("%S"))}

hours = endTime['Hours'] - startTime['Hours']
minutes = endTime['Minutes'] - startTime['Minutes']
seconds = endTime['Seconds'] - startTime['Seconds']

if seconds < 0:
    seconds += 60
    minutes -= 1

if minutes < 0:
    minutes += 60
    hours -= 1

#Final message
print(f'\nSpreadsheet built called "{fileName}".')

if minutes > 0:
    if hours > 0:
        print(f'Time to execute: {hours} hr, {minutes} min, {seconds} sec')
    else:
        print(f'Time to execute: {minutes} min, {seconds} sec')
else:
    print(f'Time to execute: {seconds} sec')