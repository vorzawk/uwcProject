import xlrd
import string
import re
from nltk.stem import WordNetLemmatizer

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
# Bad code, hardcoding the column in which the data is present
# But this will do for now!
notes = worksheet.col(0, start_rowx=0)
# if I only use the split function to identify words, "~Katie", "-Daniel", "deal," etc. will be considered as words. In this case, some
# names will be missed and some will get wrongly labeled as names.
# The solution implemented here is to ignore all punctuation and use the first unbroken chunk of letters as the word. This causes some of
# the punctuation to get missed.
# Set up the translation table to eplace all of the punctuation with spaces.
replace_punctuation = str.maketrans(string.punctuation, ' '*len(string.punctuation))
lmtr = WordNetLemmatizer()
for note in notes:
    print(note)
    wordsInNote = note.value.split()
#    print(wordsInNote)
    namesRemoved = []
    for word in wordsInNote:
#        print(word)
        if any(c.isalpha() for c in word):
            # stem the word by stripping all of its punctuation.(would it be a good idea to replace this with something else?"
            tokens = word.translate(replace_punctuation).split()
            stemWord = max(tokens, key=len)
            # if the word ends with "ed" then check if the first part is present in the word list. This is to avoid words like Worked and 
            # Checked from being labeled as names.
            #m = re.search(r'(\w+)ed', stemWord)
            #if m:
            #    # group(0) returns the entire match while group(1) returns the first subgroup.
            #    stemWord = m.group(1)
            # The lemmatizer based on WordNet is very robust in reducing the various verb forms into their roots.
            lmtWord = lmtr.lemmatize(stemWord.lower(), pos='v')
            if lmtWord != stemWord.lower():
                stemWord = lmtWord
            # Ignores common nouns which occur as plurals. The singular form is likely included in the list of common words.
            lmtWord = lmtr.lemmatize(stemWord.lower(), pos='n')
            if lmtWord != stemWord.lower():
                stemWord = lmtWord
        else:
            stemWord = ''
#        print('stem', stemWord)
        # Only consider the words in which only the first letter is capitalized while checking the list of common words. This is because
        # these are the only canditates for being names.
        if len(stemWord) > 1 and stemWord[0].isupper() and stemWord[1].islower() and stemWord.lower() not in h:
            # Suggestion: compare stemWord with word, and append the punctuation to the [NAME] token.
            namesRemoved.append('[NAME]')
            namesList.append(stemWord)
        else:
            namesRemoved.append(word)
    # namesRemoved is the list of tokens with the names redacted. This statement creates sentences from them by including spaces between
    # them.
    outputNote = ' '.join(namesRemoved)
    print(outputNote)
    print()
print('total number of words labeled as names : ', len(namesList))
print(namesList)
print()
from collections import Counter
print(Counter(namesList))


