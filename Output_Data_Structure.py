import pandas as pd
import nltk
from textblob import TextBlob

# download necessary NLTK resources
nltk.download('punkt')
nltk.download('averaged_perceptron_tagger')
nltk.download('cmudict')

# read input data from Excel file
input_df = pd.read_excel("C:/Users/ACER/Downloads/Input.xlsx")

# perform text analysis on each row
output_data = []
for _, row in input_df.iterrows():
    text = row['URL']
    blob = TextBlob(text)

    # PERCENTAGE OF COMPLEX WORDS
    words = [w for w in blob.words]
    complex_words = [w for w in words if len(w) > 2 and nltk.pos_tag([w])[0][1] in ['JJ', 'VB', 'RB', 'VBD', 'VBP', 'VBG', 'VBN', 'JJR', 'JJS', 'RBR', 'RBS']]
    percent_complex_words = (len(complex_words) / len(words)) * 100

    # FOG INDEX
    avg_sent_len = sum([len(sent.split()) for sent in blob.sentences]) / len(blob.sentences)
    fog_index = 0.4 * (avg_sent_len + percent_complex_words)

    # AVG NUMBER OF WORDS PER SENTENCE
    avg_words_per_sent = len(words) / len(blob.sentences)

    # COMPLEX WORD COUNT
    complex_word_count = len(complex_words)

    # WORD COUNT
    word_count = len(words)

    # SUBJECTIVITY AND POLARITY SCORES
    subjectivity_score = blob.sentiment.subjectivity
    polarity_score = blob.sentiment.polarity

    # POSITIVE AND NEGATIVE SCORES
    positive_score = len([w for w in words if TextBlob(w).sentiment.polarity > 0])
    negative_score = len([w for w in words if TextBlob(w).sentiment.polarity < 0])

    # SYLLABLES PER WORD
    syllables = nltk.corpus.cmudict.dict()
    syllable_count = [len(syllables.get(w.lower(), [0])) for w in words]
    syllable_per_word = sum(syllable_count) / len(syllable_count)

    # PERSONAL PRONOUNS
    personal_pronouns = ['i', 'me', 'my', 'mine', 'we', 'us', 'our', 'ours']
    personal_pronoun_count = len([w for w in words if w.lower() in personal_pronouns])

    # AVG WORD LENGTH
    avg_word_length = sum([len(w) for w in words]) / len(words)

    # append output data for this row
    output_data.append({
        'URL_ID': row['URL_ID'],
        'URL': row['URL'],
        'POSITIVE SCORE': positive_score,
        'NEGATIVE SCORE': negative_score,
        'POLARITY SCORE': polarity_score,
        'SUBJECTIVITY SCORE': subjectivity_score,
        'AVG SENTENCE LENGTH': avg_sent_len,
        'PERCENTAGE OF COMPLEX WORDS': percent_complex_words,
        'FOG INDEX': fog_index,
        'AVG NUMBER OF WORDS PER SENTENCE': avg_words_per_sent,
        'COMPLEX WORD COUNT': complex_word_count,
        'WORD COUNT': word_count,
        'SYLLABLES PER WORD': syllable_per_word,
        'PERSONAL PRONOUN COUNT': personal_pronoun_count,
})

# create a pandas DataFrame from the output data
output_df = pd.DataFrame(output_data)

# print the output DataFrame
print(output_df)

# create a pandas DataFrame from the output data
output_df = pd.DataFrame(output_data)

# write output data to Excel file
output_df.to_excel("C:/Users/ACER/OneDrive/Desktop/Professional/umm.xlsx")

