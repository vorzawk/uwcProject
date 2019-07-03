import xlrd
import string

h = set()
with open('commonWords.txt', 'r') as f:
    for line in f:
        h.add(line.rstrip())

workbook = xlrd.open_workbook('fakeSessionNotes_dummyFile.xlsx')
worksheet = workbook.sheet_by_index(0)
nrows = worksheet.nrows
ncols = worksheet.ncols

# Bad code, hardcoding the column in whichthe data is present
# But this will do for now!
notes = worksheet.col(1, start_rowx=1)
for note in notes:
    print(note)
    print()
    wordsInNote = note.value.split()
    namesRemoved = []
    for word in wordsInNote:
#        print(word)
        replace_punctuation = str.maketrans(string.punctuation, ' '*len(string.punctuation))
        stemWord = word.translate(replace_punctuation).split()[0]
#        print('stem', stemWord)
        if len(stemWord) > 1 and stemWord[0].isupper() and stemWord.lower() not in h:
            namesRemoved.append('[NAME]')
        else:
            namesRemoved.append(word)
    outputNote = ' '.join(namesRemoved)
    print(outputNote)
    print()


