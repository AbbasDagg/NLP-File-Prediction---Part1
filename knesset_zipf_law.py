import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os, sys


def is_hebrew(word):  
    try:
        if word == '':
            return False
        hebrew_letters = ['א', 'ב', 'ג', 'ד', 'ה', 'ו', 'ז', 'ח', 'ט', 'י', 'ך', 'כ', 'ל', 'ם', 'מ', 'ן', 'נ', 'ס', 'ע', 'ף', 'פ', 'ץ', 'צ', 'ק', 'ר', 'ש', 'ת']
        for letter in word:
            if letter not in hebrew_letters:
                return False
        return True
    except Exception as e:
        print(f'Exception in is_hebrew: {e}')

if __name__ == "__main__":
    try:
        if len(sys.argv) != 3:
            print('Incorrect input, please enter the folder path and the output path.')
            sys.exit(1)
            
        jsonl_file = sys.argv[1]
        output_path = sys.argv[2]

        # Read the JSONL file
        df = pd.read_json(jsonl_file, lines=True)

        # Tokenization has already been done, assuming df['text'] contains tokenized words
        
        # Count word frequencies
        frequency_dictionary = {}
        for row in df.itertuples():
            all_words = str(row[6]).split(" ")
            for word in all_words:
                if is_hebrew(word) == False:#must be hebrew word
                    continue
                if word not in frequency_dictionary.keys(): # Add it to the dictionary
                    frequency_dictionary[word]=1 
                else:
                    frequency_dictionary[word]+=1


        rank_list = [] # we sort the list
        frequency_list = []
        i=1
        for word in frequency_dictionary.keys():
            frequency_list.append(np.log2(frequency_dictionary[word]))
            rank_list.append(np.log2(i))
            i+=1
        frequency_list.sort(reverse= True)

        # Sort the frequency dictionary by values in descending order
        sorted_frequency = sorted(frequency_dictionary.items(), key=lambda x: x[1], reverse=True)

        # Extract the 10 most common words
        most_common = sorted_frequency[:10]

        # Extract the 10 least common words
        least_common = sorted_frequency[-10:]
        print(f'most common words: {most_common}')
        print(f'least common words: {least_common}')



        # Plot the Zipf's law
        plt.plot(rank_list, frequency_list)
        plt.xlabel('log(rank)')
        plt.ylabel('log(frequency)')
        plt.title('Zipf\'s Law')
        plt.savefig(output_path) # fix this 
    except Exception as e:
        print(f'Exception in main: {e}')
