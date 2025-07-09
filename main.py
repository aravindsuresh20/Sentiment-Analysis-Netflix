import pandas as pd
from textblob import TextBlob
import xlsxwriter # Required for writing Excel files with formatting

# --- Configuration ---
INPUT_CSV_FILE = "D:\original files for uplaoding\sentiment analsysis netflix/netflix_titles.csv"
OUTPUT_FILE_WITHOUT_COLOR = "D:/original files for uplaoding/sentiment analsysis netflix/netflix_titles_full_data_sentiment_numerical.csv"
OUTPUT_EXCEL_FILE_WITH_COLOR = "D:/original files for uplaoding/sentiment analsysis netflix/netflix_titles_full_data_sentiment_colored.xlsx"

# --- Sentiment Analysis Functions ---

def convert_sentiment_category(polarity):
    """
    Converts a sentiment polarity score to a categorical sentiment:
    1 for positive, -1 for negative, and 0 for neutral.
    """
    if polarity > 0:
        return 1
    elif polarity < 0:
        return -1
    else:
        return 0

def convert_polarity_to_percentage(polarity):
    """
    Converts a sentiment polarity score (range -1 to 1) to a percentage (range 0% to 100%).
    -1 becomes 0%, 0 becomes 50%, and 1 becomes 100%.
    """
    return ((polarity + 1) / 2) * 100

# --- Main Script Execution ---

if __name__ == "__main__":
    try:
        # Load the CSV file
        df = pd.read_csv(INPUT_CSV_FILE, encoding='ISO-8859-1')
        print(f"Successfully loaded '{INPUT_CSV_FILE}'.")
    except FileNotFoundError:
        print(f"Error: '{INPUT_CSV_FILE}' not found. Please ensure the file is in the same directory.")
        exit() # Exit if the input file is not found

    # Handle NaN values in the 'description' column by filling them with empty strings
    df['description'] = df['description'].fillna('')
    print("Handled NaN values in 'description' column.")

    # --- Perform Sentiment Analysis on 'description' ---

    # Calculate the raw sentiment polarity for each description
    df['description_polarity'] = df['description'].apply(lambda x: TextBlob(str(x)).sentiment.polarity)
    print("Calculated 'description_polarity'.")

    # Convert raw polarity to a numerical categorical sentiment
    df['description_sentiment_category'] = df['description_polarity'].apply(convert_sentiment_category)
    print("Calculated 'description_sentiment_category' (numerical).")

    # Convert raw polarity to a percentage
    df['description_sentiment_percentage'] = df['description_polarity'].apply(convert_polarity_to_percentage)
    print("Calculated 'description_sentiment_percentage'.")

    # --- Prepare DataFrame for Output (keeping all original columns) ---

    # Remove any existing 'rating' related sentiment columns to avoid duplicates if they exist,
    # but keep all other original columns.
    columns_to_remove_if_exist = [col for col in df.columns if 'rating_sentiment' in col or 'rating_polarity' in col]
    if columns_to_remove_if_exist:
        df = df.drop(columns=columns_to_remove_if_exist, errors='ignore')
        print(f"Removed existing rating sentiment columns: {', '.join(columns_to_remove_if_exist)}.")


    # --- Save First Output (CSV without color, all columns) ---

    # Make a copy for the first output file before any further modifications
    df_numerical_output = df.copy()

    # Save the first DataFrame (all original columns + new sentiment columns, numerical category)
    df_numerical_output.to_csv(OUTPUT_FILE_WITHOUT_COLOR, index=False)
    print(f"\nFirst dataset (all original columns + numerical sentiment) saved as '{OUTPUT_FILE_WITHOUT_COLOR}'.")

    # --- Prepare for Second Output (Excel with actual cell colors, all columns) ---

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    try:
        writer = pd.ExcelWriter(OUTPUT_EXCEL_FILE_WITH_COLOR, engine='xlsxwriter')
        df.to_excel(writer, sheet_name='Sentiment Analysis', index=False)

        # Get the XlsxWriter workbook and worksheet objects.
        workbook = writer.book
        worksheet = writer.sheets['Sentiment Analysis']

        # Get the column index for 'description_sentiment_category'
        # Pandas writes headers, so the column index starts from 0 for pd.to_excel and worksheet methods.
        header_excel = df.columns.tolist()
        try:
            category_col_idx = header_excel.index('description_sentiment_category')
        except ValueError:
            print("Error: 'description_sentiment_category' column not found for Excel output. Cannot apply conditional formatting.")
            writer.close()
            exit()

        # Define formats for colors (background color and a contrasting font color)
        red_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'}) # Light Red Fill, Dark Red Text
        green_format = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'}) # Light Green Fill, Dark Green Text
        grey_format = workbook.add_format({'bg_color': '#E0E0E0', 'font_color': '#333333'}) # Grey Fill, Dark Grey Text

        # Apply conditional formatting rules to the 'description_sentiment_category' column
        # The range starts from row 1 (the first data row after headers) to the last row of the DataFrame.
        last_row = len(df)
        worksheet.conditional_format(1, category_col_idx, last_row, category_col_idx,
                                     {'type': 'cell',    # Apply format based on cell value
                                      'criteria': '=',    # If cell value is equal to...
                                      'value': 1,         # ...1 (positive)
                                      'format': green_format}) # ...apply green format
        worksheet.conditional_format(1, category_col_idx, last_row, category_col_idx,
                                     {'type': 'cell',
                                      'criteria': '=',
                                      'value': -1,        # ...-1 (negative)
                                      'format': red_format})   # ...apply red format
        worksheet.conditional_format(1, category_col_idx, last_row, category_col_idx,
                                     {'type': 'cell',
                                      'criteria': '=',
                                      'value': 0,         # ...0 (neutral)
                                      'format': grey_format})  # ...apply grey format

        # Close the Pandas Excel writer to save the Excel file.
        writer.close()
        print(f"Second dataset (all original columns + colored sentiment) saved as '{OUTPUT_EXCEL_FILE_WITH_COLOR}'.")
        print("\nNote: This Excel file ('xlsx') supports actual cell coloring for the 'description_sentiment_category' column.")

    except ImportError:
        print("\nError: The 'xlsxwriter' library is required to create Excel files with conditional formatting.")
        print("Please install it using: pip install xlsxwriter")
    except Exception as e:
        print(f"\nAn unexpected error occurred while saving the Excel file: {e}")