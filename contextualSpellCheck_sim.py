import spacy
import contextualSpellCheck
import docx2txt
import pandas as pd
import re
from rapidfuzz import fuzz
from nltk.tokenize import PunktSentenceTokenizer
from termcolor import colored
from tkinter import filedialog
from IPython.core.display import display, HTML
display(HTML("<style>.container { width:100% !important; }</style>"))

from nltk.tokenize import sent_tokenize

file = r"C:\Users\chris\Documents\Transgola\Clients\PROJECTS\2021\401210621_TM_HS\Translation\MU COPY_HLA_P.1689_Ictiofauna_20210621 en.docx"
text = docx2txt.process(file).split("\n")
text =" ".join(text)


sents = sent_tokenize(text)

orignal_sent = []
outcome_sent = []

nlp = spacy.load('en_core_web_lg')
contextualSpellCheck.add_to_pipe(nlp)

n = 1

for s in sents:
    doc = nlp(s)
    print(f"{n}/{len(sents)} checked")
    n += 1
    if doc._.performed_spellCheck == True:
        orignal_sent.append(s) 
        outcome_sent.append(doc._.outcome_spellCheck)
        
x_list = []
y_list = []
score = []

for x,y in zip(orignal_sent, outcome_sent):
    fuzz.ratio(x, y)
    score.append(fuzz.ratio(x, y))
    x_list.append(x)
    y_list.append(y)
    
# remove consecutive blank lines

x_list1 = []  
for x in x_list:

    xn = re.sub(r'\n\s*\n', '\n\n', x)
    x_list1.append(xn)
    
y_list1 = []  
for y in y_list:

    yn = re.sub(r'\n\s*\n', '\n\n', y)
    y_list1.append(yn)
    
data_tuples = list(zip(x_list1,y_list1,score))

results = pd.DataFrame(data_tuples, columns=['X','Y', 'Score'])  

results = results.sort_values(by=['Score'], ascending=False)
results = results[results['Score'] < 100 ]

x_list3 = list(results['X'])
y_list3 = list(results['Y'])
        
    
# uncommon words

diffs = []


def find(X, Y):
    count = {}
    for word in X.split():
        count[word] = count.get(word, 0) + 1

    for word in Y.split():
        count[word] = count.get(word, 0) + 1
    return [word for word in count if count[word] == 1]



for X,Y in zip(x_list3, y_list3):
    diffs.append((find(X, Y)))
    
diffsList = [' '.join(x) for x in diffs]
results['Diffs'] = diffsList
results = results[['Score', 'X', 'Y', 'Diffs']]


resultsXlist = results['X'].tolist()
resultsYlist = results['Y'].tolist()
resultDIFFSYlist = results['Diffs'].tolist()
resultSCORElist  = results['Score'].tolist()




n = 0
while n <= len(resultsXlist) - 1:

    text1 = resultsXlist[n]  
    text2 = resultsYlist[n] 
    l1 = resultDIFFSYlist[n].split()

    
    
    formattedText1 = []
    for t in text1.split():
        if t in l1:
            formattedText1.append(colored(t,'red', attrs=['bold']))
        else: 
            formattedText1.append(t)

    
    formattedText2 = []
    for t in text2.split():
        if t in l1:
            formattedText2.append(colored(t,'red', attrs=['bold']))
        else: 
            formattedText2.append(t)
 
    print( "\n")
    print(colored(resultSCORElist[n], 'green'))
    print(colored(l1, 'blue'))
    print( "\n")
    print(" ".join(formattedText1))
    print( "\n")
    print(" ".join(formattedText2))
    print( "\n")
    print( "\n")
    
    n += 1
