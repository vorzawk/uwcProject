import xlrd
import string
import re
from nltk.stem import WordNetLemmatizer
from nltk.corpus import wordnet as wn

names = set()
with open('names.txt', 'r') as f:
    for line in f:
        names.add(line.rstrip())

patchWords = set()
with open('patch.txt', 'r') as f:
    for line in f:
        patchWords.add(line.rstrip())

commonWords = set()
with open('commonWords.txt', 'r') as f:
    for line in f:
        commonWords.add(line.rstrip())

#workbook = xlrd.open_workbook('fakeSessionNotes_dummyFile.xlsx')
workbook = xlrd.open_workbook('AllData_sample.xlsx')
worksheet = workbook.sheet_by_index(0)
nrows = worksheet.nrows
ncols = worksheet.ncols

namesList = []
# Bad code, hardcoding the column in which the data is present
# But this will do for now!
notes = worksheet.col(0, start_rowx=0)
lmtr = WordNetLemmatizer()
for note in notes:
#    print('Original Note:\n', note.value)
    print(note.value)
    wordsInNote = note.value.split()
#    print(wordsInNote)
    namesRemoved = []
    for word in wordsInNote:
        # Only consider the words in which only the first letter is capitalized while checking the list of common words. This is because
        # these are the only canditates for being names.
        m = re.search(r'([A-Z][a-z]+)(.*)', word)
        if m:
        # group(0) returns the entire match.
            stemWord = m.group(1)
        else:
            stemWord = ''
#        print('word', word)
#        print('stem', stemWord)
        synonyms = []
#        antonyms = []
        for syn in wn.synsets(stemWord):
            for l in syn.lemmas():
                synonyms.append(l.name())
#                if l.antonyms():
#                    antonyms.append(l.antonyms()[0].name())

#       print(set(synonyms))
#	print(set(antonyms))
        if (
            stemWord and
            stemWord.lower() in names or
            (
                stemWord.lower() not in commonWords and
                stemWord.lower() not in patchWords and
                not set(synonyms) & commonWords and
#            not set(antonyms) & h and
#            (not wn.synsets(stemWord) or wn.synsets(stemWord)[0].pos() not in ('r','v')) and
                (not wn.synsets(stemWord) or any([synset.pos() == 'n' for synset in wn.synsets(stemWord)])) and
#            lmtr.lemmatize(stemWord.lower(), pos='v') == stemWord.lower() and
                # Ignores common nouns which occur as plurals. The singular form is likely included in the list of common words.
                len(lmtr.lemmatize(stemWord.lower(), pos='n')) == len(stemWord)
            )
        ):
            # Suggestion: compare stemWord with word, and append the punctuation to the [NAME] token.
            namesRemoved.append('[NAME]' + m.group(2))
            namesList.append(word)
        else:
            namesRemoved.append(word)
    # namesRemoved is the list of tokens with the names redacted. This statement creates sentences from them by including spaces between
    # them.
    outputNote = ' '.join(namesRemoved)
#    print('Redacted Note: \n', outputNote)
    print(outputNote)
    print()
print('total number of words labeled as names : ', len(namesList))
print(namesList)


