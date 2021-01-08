# INIT
import codecs
from openpyxl import load_workbook
wb = load_workbook(filename='StatusyPaczek 07.01.21.xlsx', read_only=True)
out = codecs.open('output.txt', "w", "utf-8")
# CONFIG
worksheetName = 'UPS Action'
mappedField = 0
smStatus = 2
isTerminal = 3
# /CONFIG
#PRINTCONFIG
print('\n \n============================')
print('Worksheet: {0}'.format(worksheetName))
print('Mapping field index: {0}'.format(mappedField))
print('SM status field index: {0}'.format(smStatus))
print('Terminal field index: {0}'.format(isTerminal))
# /PRINTCONFIG
states = []
SM001 =[]
SM002 =[]
SM003 =[]
SM004 =[]
SM005 =[]
SM006 =[]
SM007 =[]
SM008 =[]
SM009 =[]
terminalFalse = []
terminalTrue =[]
ws = wb[worksheetName]
# /INIT
# TRANSLATE XLSX FIELDS TO STATES
for row in ws:
	if(row[isTerminal].value == None):
		terminal = False
	else:
		terminal = True
	states.append([row[mappedField].value, row[smStatus].value, terminal])
# /TRANSLATE
# SORT STATES
i = 1
while (i<len(states)):
	if(states[i][1]=="Utworzona"):
		SM001.append(states[i])
	if(states[i][1]=="Nadana"):
		SM002.append(states[i])
	if(states[i][1]=="W tranzycie"):
		SM003.append(states[i])
	if(states[i][1]=="W doreczeniu" or states[i][1]=="W doręczeniu"):
		SM004.append(states[i])
	if(states[i][1]=="Awizowana"):
		SM005.append(states[i])
	if(states[i][1]=="Doręczona"):
		SM006.append(states[i])
	if(states[i][1]=="Zwrócona"):
		SM007.append(states[i])
	if(states[i][1]=="Inny"):
		SM008.append(states[i])
	if(states[i][1]=="Odmowa przyjęcia przesyłki"):
		SM009.append(states[i])
	if(states[i][2]==True):
		terminalTrue.append(states[i])
	if(states[i][2]==False):
		terminalFalse.append(states[i])
	i=i+1
# /SORT
# CHECK STATES AND WRITE TO FILE
if(len(states)-1==len(SM001)+len(SM002)+len(SM003)+len(SM004)+len(SM005)+len(SM006)+len(SM007)+len(SM008)+len(SM009)):
	out.write('["SM001"] = new List<string>() {')
	for x in SM001:
		if SM001.index(x)==len(SM001)-1:
			out.write('"{0}"'.format(x[0]))
		else:
			out.write('"{0}", '.format(x[0]))
	out.write('},\n \n')
	out.write('["SM002"] = new List<string>() {')
	for x in SM002:
		if SM002.index(x)==len(SM002)-1:
			out.write('"{0}"'.format(x[0]))
		else:
			out.write('"{0}", '.format(x[0]))
	out.write('},\n \n')
	out.write('["SM003"] = new List<string>() {')
	for x in SM003:
		if SM003.index(x)==len(SM003)-1:
			out.write('"{0}"'.format(x[0]))
		else:
			out.write('"{0}", '.format(x[0]))
	out.write('},\n \n')
	out.write('["SM004"] = new List<string>() {')
	for x in SM004:
		if SM004.index(x)==len(SM004)-1:
			out.write('"{0}"'.format(x[0]))
		else:
			out.write('"{0}", '.format(x[0]))
	out.write('},\n \n')
	out.write('["SM005"] = new List<string>() {')
	for x in SM005:
		if SM005.index(x)==len(SM005)-1:
			out.write('"{0}"'.format(x[0]))
		else:
			out.write('"{0}", '.format(x[0]))
	out.write('},\n \n')
	out.write('["SM006"] = new List<string>() {')
	for x in SM006:
		if SM006.index(x)==len(SM006)-1:
			out.write('"{0}"'.format(x[0]))
		else:
			out.write('"{0}", '.format(x[0]))
	out.write('},\n \n')
	out.write('["SM007"] = new List<string>() {')
	for x in SM007:
		if SM007.index(x)==len(SM007)-1:
			out.write('"{0}"'.format(x[0]))
		else:
			out.write('"{0}", '.format(x[0]))
	out.write('},\n \n')
	out.write('["SM008"] = new List<string>() {')
	for x in SM008:
		if SM008.index(x)==len(SM008)-1:
			out.write('"{0}"'.format(x[0]))
		else:
			out.write('"{0}", '.format(x[0]))
	out.write('},\n \n')
	out.write('["SM009"] = new List<string>() {')
	for x in SM009:
		if SM009.index(x)==len(SM009)-1:
			out.write('"{0}"'.format(x[0]))
		else:
			out.write('"{0}", '.format(x[0]))
	out.write('},\n \n \n \n \n \n')

	print('All: {0}'.format(len(states)-1))
	print('============================')
	print('Utworzona: {0}'.format(len(SM001)))
	print('Nadana: {0}'.format(len(SM002)))
	print('W tranzycie: {0}'.format(len(SM003)))
	print('W doręczeniu: {0}'.format(len(SM004)))
	print('Awizowana: {0}'.format(len(SM005)))
	print('Doręczona: {0}'.format(len(SM006)))
	print('Zwrócona: {0}'.format(len(SM007)))
	print('Inny: {0}'.format(len(SM008)))
	print('W doręczeniu: {0}'.format(len(SM009)))
	print('============================')
	
else:
	print('Sprawdź nazwy statusów w xls')
	print(len(states), len(SM001)+len(SM002)+len(SM003)+len(SM004)+len(SM005)+len(SM006)+len(SM007)+len(SM008)+len(SM009))
# /CHECK AND WRITE
# CHECK AND WRITE TERMINALS
if(len(states)-1==len(terminalTrue)+len(terminalFalse)):
	out.write('["True"] = new List<string>() {')
	for x in terminalTrue:
		if terminalTrue.index(x)==len(terminalTrue)-1:
			out.write('"{0}"'.format(x[0]))
		else:
			out.write('"{0}", '.format(x[0]))
	out.write('},\n \n')
	out.write('["False"] = new List<string>() {')
	for x in terminalFalse:
		if terminalFalse.index(x)==len(terminalFalse)-1:
			out.write('"{0}"'.format(x[0]))
		else:
			out.write('"{0}", '.format(x[0]))
	out.write('},\n \n')

	print('Terminal: {0}'.format(len(terminalTrue)))
	print('Non terminal: {0}'.format(len(terminalFalse)))
	print('============================ \n \n')	
else:
	print('Sprawdź terminale statusów w xls')
	print(len(states), len(terminalTrue)+len(terminalFalse))
# /TERMINALS
# CLOSE FILES
wb.close()
out.close()
# /CLOSE FILES