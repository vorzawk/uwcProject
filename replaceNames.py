import xlrd
import string

h = set()
with open('commonWords.txt', 'r') as f:
    for line in f:
        h.add(line.rstrip())

#workbook = xlrd.open_workbook('fakeSessionNotes_dummyFile.xlsx')
workbook = xlrd.open_workbook('AllData_sample.xlsx')
worksheet = workbook.sheet_by_index(0)
nrows = worksheet.nrows
ncols = worksheet.ncols

namesList = []
# Bad code, hardcoding the column in whichthe data is present
# But this will do for now!
notes = worksheet.col(0, start_rowx=1)
replace_punctuation = str.maketrans(string.punctuation, ' '*len(string.punctuation))
for note in notes:
    print(note)
    wordsInNote = note.value.split()
    namesRemoved = []
    for word in wordsInNote:
#        print(word)
        if any(c.isalpha() for c in word):
            stemWord = word.translate(replace_punctuation).split()[0]
        else:
            stemWord = ''
#        print('stem', stemWord)
        if len(stemWord) > 1 and stemWord[0].isupper() and stemWord.lower() not in h:
            namesRemoved.append('[NAME]')
            namesList.append(stemWord)
        else:
            namesRemoved.append(word)
    outputNote = ' '.join(namesRemoved)
    print(outputNote)
    print()
print(namesList)


