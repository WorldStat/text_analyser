import collections
import re
import pandas as pd

def find_bigrams_and_words(input_file, output_file):
    # Read the input file
    with open(input_file, 'r') as f:
        text = f.read()

    # Split text into words
    words = re.findall(r'\b\w+\b', text)

    # Find bigrams in words
    bigrams = []
    for word in words:
        for i in range(len(word) - 1):
            bigrams.append(word[i:i+2].lower())

    # Count bigram frequencies
    bigram_frequencies = collections.Counter(bigrams)

    # Define stop words
    stop_words = {'and', 'the', 'is', 'in', 'to', 'for', 'on', 'at', 'of', 'it', 'or', 'an'}  # Add or remove as needed

    # Count word frequencies, excluding words that don't meet criteria
    word_frequencies = collections.Counter(
        word for word in words
        if word.isalpha()            # Exclude special character-only words
        and len(word) > 2            # Exclude short words
        and word.lower() not in stop_words  # Exclude stop words
    )

    # Convert bigram frequencies to a DataFrame and save to Excel
    df_bigrams = pd.DataFrame(bigram_frequencies.most_common(), columns=["Bigram", "Frequency"])
    df_words = pd.DataFrame(word_frequencies.most_common(), columns=["Word", "Frequency"])

    # Save to Excel
    with pd.ExcelWriter(output_file) as writer:
        df_bigrams.to_excel(writer, sheet_name='Bigram Frequencies', index=False)
        df_words.to_excel(writer, sheet_name='Word Frequencies', index=False)

if __name__ == "__main__":
    input_file = 'The-Feminine-Mystique.txt'
    output_file = 'The-Feminine-Mystique_output.xlsx'
    find_bigrams_and_words(input_file, output_file)
