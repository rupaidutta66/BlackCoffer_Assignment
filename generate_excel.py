import os
import pandas as pd
import requests
from bs4 import BeautifulSoup
from nltk.corpus import words, stopwords
import re
from textblob import TextBlob
import syllables
import openpyxl
from openpyxl.styles import Font



input_path = 'C:/Users/Subham/Desktop/BlackCoffee/Resources'
output_path = 'C:/Users/Subham/Desktop/BlackCoffee/Output'
positive_words_file_path = 'C:/Users/Subham/Desktop/BlackCoffee/Resources/MasterDictionary/positive-words.txt'
negative_words_file_path = 'C:/Users/Subham/Desktop/BlackCoffee/Resources/MasterDictionary/negative-words.txt'
stop_words_file_path = "C:/Users/Subham/Desktop/BlackCoffee/Resources/StopWords"
input_excel_path = os.path.join(input_path, 'Input.xlsx')
output_excel_path = os.path.join(output_path, 'Output Data Structure.xlsx')

# Read the input Excel file
input_df = pd.read_excel(input_excel_path)


def clean_and_tokenize(text):
    words = re.findall(r'\b\w+\b', text.lower())
    return [word for word in words if word not in stop_words]

def count_syllables(word):
    if word.endswith(("es", "ed")):
        return syllables.estimate(word) - 1
    return syllables.estimate(word)

def count_personal_pronouns(text):
    pronouns = re.findall(r'\b(I|we|my|ours|our|you|your|yours|they|he|she|her|him|his|me|us)\b', text, flags=re.IGNORECASE)
    return len(pronouns)

def load_words_from_file(file_path):
    with open(file_path, 'r') as file:
        words = file.read().splitlines()
    return words

positive_words = load_words_from_file(positive_words_file_path)
negative_words = load_words_from_file(negative_words_file_path)

def get_words_from_folder(folder_path):
    all_words = []
    for filename in os.listdir(folder_path):
        if filename.endswith('.txt'):  
            file_path = os.path.join(folder_path, filename)
            words = load_words_from_file(file_path)
            all_words.extend(words)
    return all_words
stop_words = get_words_from_folder(stop_words_file_path)


    # Replace these with your computed values
values = [
        ['URL_ID', 'URL', 'POSITIVE SCORE', 'NEGATIVE SCORE', 'POLARITY SCORE','SUBJECTIVITY SCORE', 'AVG SENTENCE LENGTH', 
         'PERCENTAGE OF COMPLEX WORDS', 'FOG INDEX', 'AVG NUMBER OF WORDS PER SENTENCE', 'COMPLEX WORD COUNT', 
         'WORD COUNT', 'SYLLABLE PER WORD', 'PERSONAL PRONOUNS', 'AVG WORD LENGTH']
    ]


for index, row in input_df.iterrows():
    url_id = row['URL_ID']
    url = row['URL']
    
    # Make a request to fetch the HTML content
    response = requests.get(url)
    if response.status_code == 200:
        html_content = response.text
        
        # Use BeautifulSoup to parse the HTML content
        soup = BeautifulSoup(html_content, 'html.parser')
        
        # Extract article paragraphs
        paragraphs = soup.find_all('p')  # Adjust this according to the HTML structure
        
        # Combine paragraphs into article text
        article_text = ' '.join([p.get_text() for p in paragraphs])
        
        # TextBlob analysis
        blob = TextBlob(article_text)
        '''.....................................................................................'''
        
        
        positive_score = sum(1 for word in blob.words if word.lower() in positive_words)  
        negative_score = sum(1 for word in blob.words if word.lower() in negative_words)     
        polarity_score = (positive_score - negative_score) / ((positive_score + negative_score) + 0.000001)        
        subjectivity_score = (positive_score + negative_score) / (len(blob.words) + 0.000001)

        num_words = len(blob.words)
        num_sentences = len(blob.sentences)
        
        complex_words = [word for word in blob.words if count_syllables(word) > 2]
        num_complex_words = len(complex_words)
        
        avg_sentence_length = num_words / num_sentences
        avg_words_per_sentence = num_words / num_sentences
        percentage_complex_words = num_complex_words / num_words
        fog_index = 0.4 * (avg_sentence_length + percentage_complex_words)

        # Clean and tokenize text
        cleaned_words = clean_and_tokenize(article_text)


        # Calculate the average syllable count for all words
        total_syllables = sum([count_syllables(word) for word in cleaned_words])
        average_syllables_per_word = total_syllables / len(cleaned_words)

        # Count personal pronouns
        personal_pronoun_count = count_personal_pronouns(article_text)

        # Calculate the average word length
        total_characters = sum([len(word) for word in cleaned_words])
        average_word_length = total_characters / len(cleaned_words)
        
        '''..................................................................................'''

        values.append([url_id,url,positive_score, negative_score, polarity_score, subjectivity_score, avg_sentence_length,
                       percentage_complex_words, fog_index, avg_words_per_sentence, len(complex_words),
                         len(cleaned_words),  average_syllables_per_word, personal_pronoun_count, average_word_length])
        
        print(f"Saved analysis for article {url_id}")
    else:
        print(f"Failed to fetch article {url_id}")



def create_excel_file(output_excel_path, values):
    wb = openpyxl.Workbook()
    ws = wb.active

    # Make column names bold
    bold_font = Font(bold=True)
    for col_idx in range(len(values[0])):
        cell = ws.cell(row=1, column=col_idx + 1)
        cell.value = values[0][col_idx]
        cell.font = bold_font

    # Add the remaining rows of values
    for row in values[1:]:
        ws.append(row)

    wb.save(output_excel_path)
    print(f"Excel file '{output_excel_path}' created successfully!")

create_excel_file(output_excel_path, values)