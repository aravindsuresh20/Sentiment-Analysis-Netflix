# Sentiment-Analysis-Netflix



This project provides a Python script (`main.py`) for performing sentiment analysis on descriptions of Netflix titles. It leverages the `TextBlob`
library to analyze the sentiment of each title's description and then classifies the sentiment as positive, negative, or neutral. 
The script calculates sentiment polarity, converts it to a categorical representation (1 for positive, -1 for negative, 0 for neutral), 
and also provides a sentiment percentage. The results are saved into two output files: a CSV file with numerical sentiment categories and an Excel file with conditional formatting to visually highlight positive, negative, and neutral sentiments.

## Features

* **Data Loading:** Reads Netflix titles data from a specified CSV file (e.g., `netflix_titles.csv`).
* **NaN Handling:** Automatically handles missing values in the 'description' column by filling them with empty strings.
* **Sentiment Analysis:** Calculates sentiment polarity for each 'description' using `TextBlob`.
* **Categorical Sentiment:** Converts raw polarity scores into a categorical sentiment (Positive, Negative, Neutral), represented numerically.
* **Sentiment Percentage:** Transforms polarity scores into a percentage format (0% to 100%).
* **CSV Output:** Saves the full dataset, including original columns and new numerical sentiment columns, to a CSV file.
* **Excel Output with Formatting:** Generates an Excel file where the 'description_sentiment_category' column is conditionally formatted:
* green for positive (1), red for negative (-1), and grey for neutral (0) sentiments.

## Requirements

Ensure you have the following Python dependencies installed. You can install them using `pip`:

```bash
pip install pandas textblob xlsxwriter
Key Dependencies:
•	Python 3.6+
•	pandas (for data manipulation)
•	textblob (for sentiment analysis)
•	xlsxwriter (for creating Excel files with conditional formatting)
Installation
1.	Save the Script: Download or copy the main.py file to your desired project directory.
2.	Prepare Your Dataset: Place your Netflix titles dataset as a CSV file (e.g., netflix_titles.csv) in the location specified in the script
(D:\original files for uplaoding\sentiment analsysis netflix/netflix_titles.csv), or update the INPUT_CSV_FILE variable in main.py with the correct path to your CSV file.
Usage
1.	Run the Python script: Open your terminal or command prompt, navigate to the directory where main.py is located, and execute:
Bash
python main.py
2.	Check the Outputs: Upon successful execution, two output files will be generated in the specified output directory
(D:/original files for uplaoding/sentiment analsysis netflix/):
o	netflix_titles_full_data_sentiment_numerical.csv: A CSV file containing all original data plus the new sentiment columns with numerical categories.
o	netflix_titles_full_data_sentiment_colored.xlsx: An Excel file where the 'description_sentiment_category' column is colored based on sentiment (green for positive, red for negative, grey for neutral).

##Project Structure
/your-project-directory/
│── main.py                                  # Main Python script for sentiment analysis
│── netflix_titles.csv                       # Example input dataset
│── netflix_titles_full_data_sentiment_numerical.csv  # Output CSV file
│── netflix_titles_full_data_sentiment_colored.xlsx  # Output Excel file with colored sentiment
│── README.md                                # This documentation file
