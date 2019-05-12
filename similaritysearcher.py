import graphlab
people = graphlab.SFrame.read_csv('people_mail.csv')
people['word_count'] = graphlab.text_analytics.count_words(people['Body'])
people.head()
tfidf = graphlab.text_analytics.tf_idf(people['word_count'])
people['tfidf'] = tfidf
knn_model = graphlab.nearest_neighbors.create(people,features=['tfidf'],label='name')
sahil = people[people['name'] == 'Sahil Kalra']
sahilquery = knn_model.query(sahil)
print sahilquery
sahilval = sahilquery[sahilquery['rank'] == 2]['reference_label'][0]
print sahilval
sahilreply = people[people['name'] == sahilval]['Reply']
print sahilreply

Franz = people[people['name'] == 'Franz Rottensteiner']
Franzquery = knn_model.query(Franz)
Franzval = Franzquery[Franzquery['rank'] == 2]['reference_label'][0]
Franzreply = people[people['name'] == Franzval]['Reply']
print Franzval
print Franzquery
print Franzreply
text_file = open("reply.txt", "w")

text_file.write("".join(Franzreply).replace(',', ' <br>').replace(',', ' <br>'))
text_file.write('<br>')
text_file.write('<br>')
text_file.write('<br>')

text_file.close()
