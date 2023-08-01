import os
import pandas as pd
import requests
from bs4 import BeautifulSoup
from textblob import TextBlob
import re
import math
import nltk
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize, sent_tokenize
from data_extraction import extract_article_text, save_text_to_file






def read_word_list(file_path):
    with open(file_path, 'r') as file:
        word_list = {line.strip().lower() for line in file}
    return word_list

def compute_text_analysis(article_text):

    stop_words_file = 'StopWords/Stopwords.txt'
    positive_words_file = 'MasterDict/positive-words.txt'
    negative_words_file = 'MasterDict/negative-words.txt'

    stop_words = read_word_list(stop_words_file)
    positive_words = read_word_list(positive_words_file)
    negative_words = read_word_list(negative_words_file)

    words = word_tokenize(article_text)
    cleaned_words = [word.lower() for word in words if word.isalpha() and word.lower() not in stop_words]


    # Extracting Derived Variables
    positive_score = sum(1 for word in cleaned_words if word in positive_words)
    negative_score = sum(1 for word in cleaned_words if word in negative_words)
    polarity_score = (positive_score - negative_score) / ((positive_score + negative_score) + 0.000001)
    subjectivity_score = (positive_score + negative_score) / (len(cleaned_words) + 0.000001)

    # Analysis of Readability
    sentences = sent_tokenize(article_text)
    total_words = len(cleaned_words)
    total_sentences = len(sentences)

    avg_sentence_length = total_words / total_sentences

    complex_word_count = sum(1 for word in cleaned_words if syllable_count(word) > 2)
    percentage_complex_words = (complex_word_count / total_words) * 100

    fog_index = 0.4 * (avg_sentence_length + percentage_complex_words)

    avg_words_per_sentence = total_words / total_sentences

    # Word Count
    word_count = total_words

    # Syllables per word
    total_syllables = sum(syllable_count(word) for word in cleaned_words)
    syllables_per_word = total_syllables / total_words

    # Personal Pronouns
    personal_pronouns = sum(1 for word in cleaned_words if is_personal_pronoun(word))

    # Average Word Length
    total_word_length = sum(len(word) for word in cleaned_words)
    avg_word_length = total_word_length / total_words

    return (positive_score, negative_score, polarity_score, subjectivity_score,
            avg_sentence_length, percentage_complex_words, fog_index,
            avg_words_per_sentence, complex_word_count, word_count,
            syllables_per_word, personal_pronouns, avg_word_length)

def syllable_count(word):
    word = word.lower()
    if len(word) <= 3:
        return 1
    return len(re.findall(r'[aeiouy]{1,2}', word))

def is_personal_pronoun(word):
    pronouns = {'i', 'you', 'he', 'she', 'it', 'we', 'they', 'me', 'us', 'him', 'her', 'them'}
    return word.lower() in pronouns

def main():
    input_file = 'Input.xlsx'
    df = pd.read_excel(input_file)

    # Initialize columns for computed variables in the DataFrame
    df['POSITIVE SCORE'] = None
    df['NEGATIVE SCORE'] = None
    df['POLARITY SCORE'] = None
    df['SUBJECTIVITY SCORE'] = None
    df['AVG SENTENCE LENGTH'] = None
    df['PERCENTAGE OF COMPLEX WORDS'] = None
    df['FOG INDEX'] = None
    df['AVG NUMBER OF WORDS PER SENTENCE'] = None
    df['COMPLEX WORD COUNT'] = None
    df['WORD COUNT'] = None
    df['SYLLABLE PER WORD'] = None
    df['PERSONAL PRONOUNS'] = None
    df['AVG WORD LENGTH'] = None

    for index, row in df.iterrows():
        url_id = row['URL_ID']
        url = row['URL']

        article_data = extract_article_text(url)
        if article_data is not None:
            article_title, article_text = article_data

            # Compute text analysis and variables
            (positive_score, negative_score, polarity_score, subjectivity_score,
             avg_sentence_length, percentage_complex_words, fog_index,
             avg_words_per_sentence, complex_word_count, word_count,
             syllables_per_word, personal_pronouns, avg_word_length) = compute_text_analysis(article_text)

            # Save the article text to a text file
            save_text_to_file(url_id, article_title, article_text)

            # Update the DataFrame with the computed variables
            df.loc[index, 'POSITIVE SCORE'] = positive_score
            df.loc[index, 'NEGATIVE SCORE'] = negative_score
            df.loc[index, 'POLARITY SCORE'] = polarity_score
            df.loc[index, 'SUBJECTIVITY SCORE'] = subjectivity_score
            df.loc[index, 'AVG SENTENCE LENGTH'] = avg_sentence_length
            df.loc[index, 'PERCENTAGE OF COMPLEX WORDS'] = percentage_complex_words
            df.loc[index, 'FOG INDEX'] = fog_index
            df.loc[index, 'AVG NUMBER OF WORDS PER SENTENCE'] = avg_words_per_sentence
            df.loc[index, 'COMPLEX WORD COUNT'] = complex_word_count
            df.loc[index, 'WORD COUNT'] = word_count
            df.loc[index, 'SYLLABLE PER WORD'] = syllables_per_word
            df.loc[index, 'PERSONAL PRONOUNS'] = personal_pronouns
            df.loc[index, 'AVG WORD LENGTH'] = avg_word_length

            print(f"Article {url_id} extracted and analyzed successfully.")
        else:
            # Create a text file with an error message for URLs that couldn't be opened
            error_message = f"Failed to open URL: {url}"
            save_text_to_file(url_id, None, None, error_message)
            print(f"Article {url_id} couldn't be opened. Error message saved to file.")


    output_file = 'Output.xlsx'
    df.to_excel(output_file, index=False)

if __name__ == "__main__":
    main()
