import pandas as pd
from datetime import datetime,date
import xlsxwriter
from urllib.request import Request, urlopen
import nltk
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize
from nltk.tokenize import sent_tokenize
from nltk.stem import PorterStemmer
import string
import re
import time

# starting a timer
start = time.time()

# cleaning html code
def cleanhtml(raw_html):
	cleanr = re.compile('<.*?>')
	cleantext = re.sub(cleanr, '', raw_html)
	return cleantext

file =('cik_list.xlsx') # file path
cik_list = pd.read_excel(file, engine ='openpyxl',nrows=152) # reading excel file
output_df = pd.DataFrame(cik_list) # dataframe which we will our output
# making empty lists for other variables
constraining_words_whole_reports=[]
positive_scores=[]
negative_scores=[]
polarity_scores= []
average_sentence_lengths= []
percentage_of_complex_wordss=[]
fog_indexs= []
complex_word_counts= []
word_counts= []
uncertainty_scores= []
constraining_scores= []
positive_word_proportions=[]
negative_word_proportions= []
uncertainty_word_proportions= []
constraining_word_proportions=[]

 # using output_df.index to iterate through dataframe
for ind in output_df.index: 
          # fetching the details from the link
          financial_report_link_req = Request("https://www.sec.gov/Archives/"+output_df["SECFNAME"][ind], headers={'User-Agent': 'XYZ/3.0'})
          financial_report_links = urlopen(financial_report_link_req).read()
          financial_report = str(financial_report_links, 'UTF-8')    
          #converting bytes to string
          #financial_report = driver.page_source
          # removing html elements
          financial_report = cleanhtml(financial_report)
          # using sent_tokenize from nltk to find no of sentences
          no_of_sent = len(sent_tokenize(financial_report))
          # using word_tokenize to separate words
          words = word_tokenize(financial_report)
          # string.punctuation contains all punctuations
          punctuation = list(string.punctuation) 
          # removing punctuations
          for word in words:                         
                  if word in punctuation:
                          words.remove(word)
          # removing stopwords
          for word in words:
                  if word in stopwords.words('english'):
                          words.remove(word)
          # remove html element nbsp and digits/numbers
          for word in words:
                  if word == "nbsp" or word.isdigit():
                          words.remove(word)
          # actual no of words
          no_of_words = len(words)
          print(no_of_words)
          # reading positive words
          positive_words = pd.read_excel('LoughranMcDonald_SentimentWordLists_2018.xlsx', sheet_name='Positive',engine ='openpyxl')
          positive=[]
          # adding all positive words from url in a list
          for word in words:
                  word=word.upper()
                  if word in positive_words.values:
                          positive.append(word)
          # calculating positive score
          positive_score=len(positive)
          # reading negative words
          negative_words = pd.read_excel('LoughranMcDonald_SentimentWordLists_2018.xlsx', sheet_name='Negative',engine ='openpyxl')
          negative=[]
          # adding all negative words from url in a list
          for word in words:
                  word=word.upper()
                  if word in negative_words.values:
                          negative.append(word)
          # calculating negative score
          negative_score=len(negative)
          # calculating polarity score
          polarity_score=(positive_score - negative_score)/ ((positive_score + negative_score) + 0.000001) * -1
          # calculating average sentence length
          average_sent_length=no_of_words/no_of_sent
          # stemming words list  
          ps = PorterStemmer()
          for word in words:
                  words[words.index(word)]=ps.stem(word)
          complex_words=[]
          # syllable count
          # if text contain more than 2 syllables then declaring it as complex
          for i in words:
                  no_of_vowel = i.count('a')+i.count('e')+i.count('i')+i.count('o')+i.count('u')
                  if no_of_vowel > 2:
                          complex_words.append(i)
          complex_word_count=len(complex_words)
          # calculating percentage_of_complex_words length
          percentage_of_complex_words = complex_word_count / no_of_words
          # calculating fog_index
          fog_index =  0.4 * (average_sent_length + percentage_of_complex_words)
          # adding all uncertainty words from url in a list
          uncertainty_words = pd.read_excel('LoughranMcDonald_SentimentWordLists_2018.xlsx', sheet_name='Uncertainty',engine ='openpyxl')
          uncertainty=[]
          for word in words:
                  word=word.upper()
                  if word in uncertainty_words.values:
                          uncertainty.append(word)
          # calculating uncertainty score
          uncertainty_score=len(uncertainty)
          # adding all constraining words from url in a list
          constraining_words = pd.read_excel('LoughranMcDonald_SentimentWordLists_2018.xlsx', sheet_name='Constraining',engine ='openpyxl')
          constraining=[]
          for word in words:
                  word=word.upper()
                  if word in constraining_words.values:
                          constraining.append(word)
          # calculating percentage_of_complex_words length
          constraining_score=len(constraining)
          # calculating remaining variables
          positive_word_proportion = positive_score / no_of_words
          negative_word_proportion = negative_score / no_of_words
          uncertainty_word_proportion = uncertainty_score / no_of_words
          constraining_word_proportion = constraining_score / no_of_words
          constraining_words_whole_report = len(constraining)
          positive_scores.append(positive_score)
          negative_scores.append(negative_score)
          polarity_scores.append(float(format(polarity_score,'.4f')))
          average_sentence_lengths.append(float(format(average_sent_length,'.4f')))
          # appending variables value to respective lists
          percentage_of_complex_wordss.append(float(format(percentage_of_complex_words,'.4f')))
          fog_indexs.append(float(format(fog_index,'.4f')))
          complex_word_counts.append(complex_word_count)
          word_counts.append(no_of_words)
          uncertainty_scores.append(uncertainty_score)
          constraining_scores.append(constraining_score)
          positive_word_proportions.append(float(format(positive_word_proportion,'.4f')))
          negative_word_proportions.append(float(format(negative_word_proportion,'.4f')))
          uncertainty_word_proportions.append(float(format(uncertainty_word_proportion,'.4f')))
          constraining_word_proportions.append(float(format(constraining_word_proportion,'.4f')))
          constraining_words_whole_reports.append(constraining_words_whole_report)
          print(positive_scores,negative_scores,polarity_scores,average_sentence_lengths,percentage_of_complex_wordss,
                fog_indexs,complex_word_counts,word_counts,uncertainty_scores,constraining_scores,positive_word_proportions,
                negative_word_proportions,uncertainty_word_proportions,constraining_word_proportions,constraining_words_whole_reports)


# assigning values 
output_df['positive_score'] = positive_scores
output_df['negative_score'] = negative_scores
output_df['polarity_score'] = polarity_scores
output_df['average_sentence_length'] = average_sentence_lengths
output_df['percentage_of_complex_words'] = percentage_of_complex_wordss
output_df['fog_index'] = fog_indexs
output_df['complex_word_count'] = complex_word_counts
output_df['word_count'] = word_counts
output_df['uncertainty_score'] = uncertainty_scores
output_df['constraining_score'] = constraining_scores
output_df['positive_word_proportion'] = positive_word_proportions
output_df['negative_word_proportion'] = negative_word_proportions
output_df['uncertainty_word_proportion'] = uncertainty_word_proportions
output_df['constraining_word_proportion'] = constraining_word_proportions
output_df['constraining_words_whole_report'] = constraining_words_whole_reports

# converting dataframe to excel
output_df.to_excel('output.xlsx',index=False)

# ending the timer
end = time.time()

print(f"Runtime of the program is {end - start}")
