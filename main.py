import pandas as pd
import requests
from bs4 import BeautifulSoup
from textblob import TextBlob
from textstat import flesch_reading_ease, syllable_count, textstat
import csv

class URLAnalyzer:
    def __init__(self, excel_file_path, output_csv_path):
        # Initialize the URLAnalyzer with paths to Excel file and output CSV file
        self.excel_file_path = excel_file_path
        self.output_csv_path = output_csv_path

    def extract_content_and_analyze(self, url):
        # Extract content from URL and analyze
        try:
            response = requests.get(url)
            response.raise_for_status()  # Raise exception for HTTP errors
            if response.status_code == 200:
                soup = BeautifulSoup(response.text, 'html.parser')
                # Find content divs based on specific classes
                div_content = soup.find('div', class_='td-post-content tagdiv-type')
                if div_content is None:
                    div_content = soup.find('div', class_='tdb-block-inner td-fix-index')

                h1_tag = soup.find('h1', class_='entry-title')
                if h1_tag is None:
                    h1_tag = soup.find('h1', class_='tdb-title-text')

                extracted_text = ""

                if h1_tag:
                    # Extract title and append to extracted_text
                    extracted_text += "\n\nTitle: " + h1_tag.text.strip()
                
                if div_content:
                    extracted_text += div_content.get_text(separator='\n').strip()
                return self.analyze_text(extracted_text)
            else:
                # Return a list of zeros if response status is not 200
                return [0] * 15  # Adjust the list size to accommodate new metrics
        except Exception as e:
            print(f"Error occurred while extracting content from {url}: {e}")
            return [0] * 15  # Adjust the list size to accommodate new metrics

    def analyze_text(self, text):
        # Analyze text content
        try:
            blob = TextBlob(text)
            sentences = blob.sentences
            words = blob.words

            positive_score = sum(sentence.sentiment.polarity for sentence in sentences if sentence.sentiment.polarity > 0)
            negative_score = sum(sentence.sentiment.polarity for sentence in sentences if sentence.sentiment.polarity < 0)

            polarity_score = blob.sentiment.polarity
            subjectivity_score = blob.sentiment.subjectivity

            avg_sentence_length = len(words) / len(sentences) if len(sentences) > 0 else 0

            complex_words = [word for word in words if textstat.syllable_count(word) > 2]
            complex_word_count = len(complex_words)
            percentage_complex_words = (complex_word_count / len(words)) * 100
            fog_index = 0.4 * (avg_sentence_length + percentage_complex_words)

            avg_words_per_sentence = len(words) / len(sentences) if len(sentences) > 0 else 0
            word_count = len(words)
            avg_word_length = sum(len(word) for word in words) / len(words) if len(words) > 0 else 0

            syllables_per_word = sum(syllable_count(word) for word in words) / len(words) if len(words) > 0 else 0
            personal_pronouns = sum(word.lower() in ["i", "me", "my", "mine", "myself", "we", "us", "our", "ours", "ourselves"] for word in words)

            return [positive_score, negative_score, polarity_score, subjectivity_score, avg_sentence_length,
                    percentage_complex_words, fog_index, avg_words_per_sentence, complex_word_count, word_count,
                    avg_word_length, syllables_per_word, personal_pronouns]
        except Exception as e:
            print(f"Error occurred while analyzing text: {e}")
            return [0] * 13  # Adjust the list size to accommodate new metrics

    def analyze_urls(self):
        # Analyze URLs listed in the Excel file
        try:
            df = pd.read_excel(self.excel_file_path)
            analysis_results = []

            for i in range(len(df["URL_ID"])):
                url_id = df["URL_ID"][i]
                url = df["URL"][i]
                #print(url_id)
                analysis_result = self.extract_content_and_analyze(url)
                analysis_results.append([url_id, url] + analysis_result)

            self.write_to_csv(analysis_results)
            print("Analysis complete. Results written to:", self.output_csv_path)
        except Exception as e:
            print("Error occurred during URL analysis:", e)

    def write_to_csv(self, data):
        # Write analysis results to a CSV file
        try:
            columns = ["URL_ID", "URL", "POSITIVE SCORE", "NEGATIVE SCORE", "POLARITY SCORE", "SUBJECTIVITY SCORE",
                       "AVG SENTENCE LENGTH", "PERCENTAGE OF COMPLEX WORDS", "FOG INDEX",
                       "AVG NUMBER OF WORDS PER SENTENCE",
                       "COMPLEX WORD COUNT", "WORD COUNT", "AVG WORD LENGTH", "SYLLABLES PER WORD", "PERSONAL PRONOUNS"]

            with open(self.output_csv_path, 'w', newline='') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(columns)
                writer.writerows(data)
        except Exception as e:
            print("Error occurred while writing to CSV:", e)


# Example usage:
if __name__ == "__main__":
    # Define input and output file paths
    excel_file_path = "./project/files/input.xlsx"
    output_csv_path = "./output.csv"
    # Instantiate URLAnalyzer and perform analysis
    analyzer = URLAnalyzer(excel_file_path, output_csv_path)
    analyzer.analyze_urls()
