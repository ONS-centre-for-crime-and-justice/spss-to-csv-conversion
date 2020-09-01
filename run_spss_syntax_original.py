# This script is not intended to be directly run. As the python modules bundled with SPSS 24 are only compatible with
# python 2.7 and 3.4 and our machines are set up with 3.6, this script must be run with the 3.4 python interpreter that
# comes bundled with SPSS.


import SpssClient
import spssaux
import spss
import csv
import re

SpssClient.StartClient()

# User defined variables
input_file = r'user_input_1'
output_file = r'user_input_2'
output_var_vals_file = r'user_input_3'
output_var_info_file = r'user_input_4'

# Load sav-file.
spssaux.OpenDataFile(input_file)

# Count number of columns
varCount = spss.GetVariableCount()
caseCount = spss.GetCaseCount()
print('There are %d variables in this file' % varCount)
print('There are %d cases. Please check this matches number of cases in the output file' % spss.GetCaseCount())

# Clean file: only string columns
for ind in range(varCount):
    varName = spss.GetVariableName(ind)
    if spss.GetVariableType(ind) > 0:
        print("String Variable: %s" % varName)
        for ascii_ind in range(33):
            spss.Submit('compute %s = REPLACE(%s, STRING(%d,PIB), " ")' % (varName, varName, ascii_ind))
        spss.Submit('EXECUTE.')

# Save to csv
spss.Submit(r"""
    SAVE TRANSLATE OUTFILE= '%s'
    /TYPE=CSV /ENCODING='Locale'
    /MAP /REPLACE /FIELDNAMES
    /CELLS=VALUES.
    """ % (output_file))

SpssClient.StopClient()

#remove whitespace for missing values
new_rows_list = []
    
file = open(r'user_input_2', 'r+', newline='')
for row in csv.reader(file):
    new_row = [re.sub('^\s*$', '', item) for item in row]
    new_rows_list.append(new_row)
file.close() 

file = open(r'user_input_2', 'w', newline='')
writer = csv.writer(file)
writer.writerows(new_rows_list)
file.close()   

SpssClient.StartClient()

spss.Submit(r"""oms SELECT tables /if subtypes=['Variable Values']
    /destination format=XLS outfile='%s'.
    oms SELECT tables /if subtypes='Variable Information'
    /destination format=XLS outfile='%s'.
    display dict.
    omsend.""" % (output_var_vals_file, output_var_info_file))

# SpssClient.Exit()
SpssClient.StopClient()

infile = open(r'user_input_2', 'r')

csv_varCount = len(next(csv.reader(infile)))
csv_caseCount = sum(1 for row in infile)

infile.close()

assert csv_varCount == varCount, 'Number of variables in csv does not match SPSS .sav'
assert csv_caseCount == caseCount, 'Number of cases in csv does not match SPSS .sav'

print('There are %d variables in the .sav and %d variables in the .csv' % (varCount, csv_varCount))
print('There are %d cases in the .sav and %d cases in the .csv' % (caseCount, csv_caseCount))



