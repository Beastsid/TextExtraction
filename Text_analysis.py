# -*- coding: utf-8 -*-
"""
Created on Sun Apr 23 23:02:33 2023

@author: ACER
"""
import pandas as pd
import os
import nltk
from textblob import TextBlob
nltk.download('averaged_perceptron_tagger')
from textblob import TextBlob

# read in the text file
with open('C:/Users/ACER/Downloads/Output Data Structure.xlsx', 'r', encoding='iso-8859-1') as f:
    text = f.read()

# create a TextBlob object
blob = TextBlob(text)

# POSITIVE SCORE
pos_score = sum([s.sentiment.polarity for s in blob.sentences if s.sentiment.polarity > 0])

# NEGATIVE SCORE
neg_score = sum([s.sentiment.polarity for s in blob.sentences if s.sentiment.polarity < 0])

# POLARITY SCORE
polarity_score = blob.sentiment.polarity

# SUBJECTIVITY SCORE
subjectivity_score = blob.sentiment.subjectivity

# AVG SENTENCE LENGTH
avg_sent_len = sum([len(s.split()) for s in blob.sentences]) / len(blob.sentences)

# PERCENTAGE OF COMPLEX WORDS
words = [w for w in blob.words]
complex_words = [w for w in words if len(w) > 2 and nltk.pos_tag([w])[0][1] in ['JJ', 'VB', 'RB', 'VBD', 'VBP', 'VBG', 'VBN', 'JJR', 'JJS', 'RBR', 'RBS']]
percent_complex_words = (len(complex_words) / len(words)) * 100

# FOG INDEX
fog_index = 0.4 * (avg_sent_len + percent_complex_words)

# AVG NUMBER OF WORDS PER SENTENCE
avg_words_per_sent = len(words) / len(blob.sentences)

# COMPLEX WORD COUNT
complex_word_count = len(complex_words)

# WORD COUNT
word_count = len(words)

# SYLLABLE PER WORD
def count_syllables(word):
    return len([v for v in word if v.lower() in 'aeiouy'])
syllables = [count_syllables(w) for w in words]
syllables_per_word = sum(syllables) / len(words)

# PERSONAL PRONOUNS
personal_pronouns = ['I', 'me', 'my', 'mine', 'myself', 'we', 'us', 'our', 'ours', 'ourselves']
personal_pronoun_count = sum([1 for w in words if w.lower() in personal_pronouns])

# AVG WORD LENGTH
avg_word_len = sum([len(w) for w in words]) / len(words)