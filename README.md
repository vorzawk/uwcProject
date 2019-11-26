# Notes Anonymizer
I completed this project while working as a Writing Consultant at the [Texas A&M University Writing Center](https://writingcenter.tamu.edu/). We collected a dataset of session notes from our writing center as well as others in the region. Our goal was to analyze the data to glean insights which would help us improve the effectiveness of our sessions. The first step was to process the data so that the dataset did not compromise the identities of the students, and this program accomplishes exactly this.

The program reads through every word in the session notes and decides whether the given word is a name or not. 
The words identified as names are replaced by the [NAME] token.

The program first detects capitalized words in the notes, and then queries the WordNet database for them. The program uses the 
information obtained from the [WordNet](https://wordnet.princeton.edu/) database to decide whether the given word is a name or not. 
The program also allows the user to manually correct the program output, by pre-specifying certain words as names (or common words), 
which overrides the program output, and can be used to correct its mistakes.
