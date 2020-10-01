import numpy as np
import pandas as pd
# from docx import Document
import docx2txt

data = [pd.read_excel('Eamcet Key-converted.xlsx', sheet_name = 1, usecols = [1,2]),
pd.read_excel('Eamcet Key-converted.xlsx', sheet_name = 1, usecols = [4,5]),
pd.read_excel('Eamcet Key-converted.xlsx', sheet_name = 1, usecols = [7,8]),
pd.read_excel('Eamcet Key-converted.xlsx', sheet_name = 1, usecols = [10,11])]
questions = {}
for dat in data:
    for index, row in dat.iterrows():
        questions[row[0]] = row[1]
# print(questions)


# document = Document('Eamcet Sravan Response Sheet-converted.docx')

# fullText = []
# for para in document.paragraphs:
#     fullText.append(para.text)
# print(fullText)

ques_text = docx2txt.process("filename.docx")
ques_list = ques_text.split()
# print(ques_list)

q_a = []
for i in range(len(ques_list)):
    if('Question' in ques_list[i]):
        if(ques_list[i+1] == 'Type'):
            continue
        q = int(ques_list[i+3])
        
        ans = int(ques_list[i+7])
        q_a.append([q,ans])

# print(q_a)
score = 0
for i in q_a:
    if(i[1] == questions[i[0]]):
        score += 1
    print(i[1],i[0],questions[i[0]])
print(score)
