import spacy

# Load English tokenizer, tagger, parser, NER and word vectors
nlp = spacy.load("en_core_web_sm")

# Process whole documents
f = open("reply.txt", "r")
text = f.read()

print(text)
doc = nlp(text)

# Analyze syntax
print("Noun phrases:", [chunk.text for chunk in doc.noun_chunks])
print("Verbs:", [token.lemma_ for token in doc if token.pos_ == "VERB"])
Noun_phrases = [chunk.text for chunk in doc.noun_chunks]
Verbs = [token.lemma_ for token in doc if token.pos_ == "VERB"]
replacelist = list(set(text.split(' ')).intersection(set(Noun_phrases)))
print(replacelist)
for i in replacelist:
    if len(i)> 1:
        text = text.replace(i,'{ }')
replacelist = list(set(text.split(' ')).intersection(set(Verbs)))
for i in replacelist:
    if len(i)> 1:
        text = text.replace(i,'{ }')
print(text)

text_file = open("reply.txt", "a")

text_file.write('<br>')
text_file.write('<br>')
text_file.write('<br>')
text_file.write('<br>DraftedText::<br>')
text_file.write(text.replace('.','.<br>').replace(',',',<br>'))

f = open("reply1.txt", "r")
text = f.read()
print(text)



# from markovipy import MarkoviPy
# obj = MarkoviPy(r"output.txt", 3)
# print(obj.generate_sentence())

import win32com.client

o = win32com.client.Dispatch("Outlook.Application")

Msg = o.CreateItem(0)
Msg.Importance = 0
Msg.Subject = 'SUBJECT'
Msg.HTMLBody = text


Msg.To = 'markov@outlook.com'
#Msg.BCC = STRING_CONTAINING_BCC

Msg.SentOnBehalfOfName = 'varada@outlook.com'
Msg.ReadReceiptRequested = True
Msg.OriginatorDeliveryReportRequested = True

Msg.save()