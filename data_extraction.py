import os
import pandas as pd
import requests
from bs4 import BeautifulSoup

def extract_article_text(url):
    try:
        # Send an HTTP request to get the webpage content
        response = requests.get(url)
        response.raise_for_status()

        # Parse the HTML content using BeautifulSoup
        soup = BeautifulSoup(response.content, 'html.parser')

        # Find the article title and text elements using their HTML tags (modify as needed)
        article_title = soup.find('h1').text.strip()
        article_text = '\n'.join([p.text.strip() for p in soup.find_all('p')])

        return article_title, article_text
    except requests.exceptions.RequestException as e:
        print(f"Failed to fetch data from {url}. Error: {e}")
        return None
    except Exception as e:
        print(f"An error occurred while processing {url}. Error: {e}")
        return None

def save_text_to_file(url_id, article_title, article_text, error_message=None):
    # Create a directory to save the text files if it doesn't exist
    if not os.path.exists('article_texts'):
        os.makedirs('article_texts')

    # Save the article text or error message to a text file
    file_name = f'article_texts/{url_id}.txt'
    with open(file_name, 'w', encoding='utf-8') as file:
        if error_message:
            file.write(error_message)
        else:
            file.write(f'{article_title}\n\n{article_text}')

def main():
    # Read the input file
    input_file = 'Input.xlsx'
    df = pd.read_excel(input_file)

    for index, row in df.iterrows():
        url_id = row['URL_ID']
        url = row['URL']

        # Extract article text
        article_data = extract_article_text(url)
        if article_data is not None:
            article_title, article_text = article_data

            # Save the article text to a text file
            save_text_to_file(url_id, article_title, article_text)
            print(f"Article {url_id} extracted and saved successfully.")
        else:
            # Create a text file with an error message for URLs that couldn't be opened
            error_message = f"Failed to open URL: {url}"
            save_text_to_file(url_id, None, None, error_message)
            print(f"Article {url_id} couldn't be opened. Error message saved to file.")

if __name__ == "__main__":
    main()
