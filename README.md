# Python Bigram and Word Frequency Calculator

This Python script reads a text file, calculates the most frequent two-letter combinations (bigrams) and words within the text, and saves these frequencies to different tabs in an Excel file. It provides a way to uncover patterns in text that can be used for enhancing your calligraphy skills!

## How It Works

1. Reads an input file (default is 'input.txt').
2. Uses a regular expression to find all words in the text.
3. Extracts all two-letter combinations (bigrams) from each word, treating upper- and lowercase letters as the same.
4. Counts the frequency of each bigram and word, and sorts them by their frequencies.
5. Only words that meet certain criteria are counted:
   - They consist of alphabetic characters only.
   - They are longer than 2 characters.
   - They are not in the defined list of stop words (common words like "and", "the", "is", etc.).
6. Saves the frequencies to different tabs in an output Excel file (default is 'output.xlsx').

By default, the script reads from 'input.txt' and writes to 'output.xlsx'. You can modify the input_file and output_file variables in the script to read from and write to different files.

## Dependencies
This script depends on the pandas and openpyxl libraries. You can install these with pip:

shell
Copy code
pip install pandas openpyxl
Note
This script only counts bigrams within each word (i.e., 'this is a test' only includes the bigrams 'th', 'hi', 'is', 'te', and 'st'). If you want to count bigrams across word boundaries, you will need to modify the script.


## Usage

To run the script, use the following command:

```shell
python3 bigram_word_frequency_calculator.py
