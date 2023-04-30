# -*- coding: utf-8 -*-
"""
Created on Sun Apr 23 20:18:02 2023

@author: ACER
"""
#Data Extraction 
#Importing Necessary Libraries
import requests
from bs4 import BeautifulSoup
import pandas as pd
import os
import nltk
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize, sent_tokenize
import string
from textblob import TextBlob

# Step 1: Extract article text and save in text file
df = pd.read_excel("C:/Users/ACER/Downloads/Input.xlsx", engine="openpyxl")
output_dir = "C:/Users/ACER/Downloads/TextFiles/"
print(df)

if not os.path.exists(output_dir):
    os.makedirs(output_dir)

#Extract article text and save in text file
for index, row in df.iterrows():
    url = row["URL"]
    url_id = row["URL_ID"]
    response = requests.get(url)
    soup = BeautifulSoup(response.content, "html.parser")
    article = soup.find("article")
    if article:
        article_title = article.find("h1").get_text().strip()
        article_text = "\n\n".join([p.get_text().strip() for p in article.find_all("p")])
    else:
        article_title = ""
        article_text = ""
    with open(f"{output_dir}/{url_id}.txt", "w", encoding="utf-8") as f:
        f.write(f"{article_title}\n\n{article_text}")





#DATA ANALYSIS
# extracted text files and load into a dictionary
articles = {}
for filename in os.listdir(output_dir):
    if filename.endswith(".txt"):
        with open(os.path.join(output_dir, filename), "r", encoding="utf-8") as f:
            articles[filename[:-4]] = f.read()
print(articles)

# Preprocess text data
nltk.download("stopwords")
nltk.download("punkt")
stop_words = set(stopwords.words("english"))
punctuations = set(string.punctuation)

def preprocess(text):
    words = word_tokenize(text.lower())
    words = [w for w in words if w not in stop_words and w not in punctuations]
    return words

# Compute variables
output = []
for url_id, text in articles.items():
    # Compute word count
    words = preprocess(text)
    word_count = len(words)
    
    # Compute sentence count
    sentences = sent_tokenize(text)
    sentence_count = len(sentences)
    
    # Compute average sentence length
    words_per_sentence = word_count / sentence_count if sentence_count > 0 else 0
    
    # Compute sentiment scores
    blob = TextBlob(text)
    polarity_score = blob.sentiment.polarity
    subjectivity_score = blob.sentiment.subjectivity
    
    output.append([url_id, word_count, sentence_count, words_per_sentence, polarity_score, subjectivity_score])


# Save output in Excel file
df_output = pd.DataFrame(output, columns=["URL_ID", "Word_Count", "Sentence_Count", "Avg_Sentence_Length", "Polarity_Score", "Subjectivity_Score"])
with pd.ExcelWriter("C:/Users/ACER/Downloads/Output Data Structure (1).xlsx", engine="openpyxl") as writer:
    df_output.to_excel(writer, index=False)
print(df_output)







