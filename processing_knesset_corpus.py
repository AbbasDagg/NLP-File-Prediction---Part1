from docx import Document
import zipfile
import os
import pandas
import re
#from tqdm import tqdm
import sys

def get_docx(zip_file_path):
    try:
        info = []

        # Open the ZIP file
        with zipfile.ZipFile(zip_file_path, 'r') as zipf:
            # Loop through the files in the ZIP
            for filename in zipf.namelist():
                if filename.endswith('.docx'):
                    attributes = filename.split('_')  # Get and Seperate the attributes 

                    if attributes[1] == 'ptv':  # committee or plenary
                        type = 'committee'
                    elif attributes[1] == 'ptm':
                        type = 'plenary'
                    else:
                        type = '-1'

                    # Read the DOCX file from within the ZIP
                    with zipf.open(filename) as file:
                        text = Document(file)
                    # Next item in the list

                    info_item = {
                        'file_name': filename,
                        'knesset number': int(attributes[0]),
                        'type': type,
                        'text': text,
                        'file_number': attributes[2].replace('.docx', '')
                    }
                    info.append(info_item)

        return info
    except Exception as e:
        print(f"Exception in get_docx: {e}")
        return None

# Function to the number after we find "הישיבה ה" or "'פרוטוקול מס"
def get_next_word(text, position):
    # Find the start of the next word
    word_start = position


    while word_start < len(text) and text[word_start].isspace(): # Skip spaces
        word_start += 1

    # If reached the end of the text, return empty
    if word_start >= len(text):
        return "-1"

    # Find the end of the next word
    word_end = word_start
    while word_end < len(text) and not text[word_end].isspace():
        word_end += 1

    # Return the next continuous word
    return text[word_start:word_end]


if __name__ == "__main__":
    try:
        print(sys.argv)
        if len(sys.argv) !=3:
            print('Incorrect input, please enter the folder path and the output path.')
            sys.exit(1)

        # Add Check if folder is valid


        folder_path = sys.argv[1]
        output_path = sys.argv[2]
        info = get_docx(folder_path)
        cpy = info.copy()
        print(info[0])
        target_words = ["הישיבה ה", "פרוטוקול מס'"]  # Search for the protocal number
        jsonl_data = []

        for doc_num, doc in enumerate(info):
            knesset_number = doc['knesset number']
            protocol_type = doc['type']
            file_number = doc['file_number']

            for par in doc['text'].paragraphs:
                text = par.text.strip() # Remove leading and trailing spaces

                if text.startswith('<') or text.startswith('>'): # Sometimes the text starts with < and ends with >, its probably caused by the conversion from doc to docx so we remove it
                    text = text[1:-1] 

                if target_words[0] in text:
                    position = text.find(target_words[0])
                    next_word = get_next_word(text, position + len(target_words[0]))
                    print(f"Found in doc {doc_num}: {text}, NEXT {next_word}")
                    protocol_number = next_word
                    break
                
                if target_words[1] in text:
                    position = text.find(target_words[1])
                    next_word = get_next_word(text, position + len(target_words[1]))
                    print(f"Found in doc {doc_num}: {text}, NEXT {next_word}")
                    protocol_number = next_word
                    break
                    
            # Append the data to the jsonl_data list
            jsonl_data.append({
                'knesset_number': knesset_number,
                'protocol_type': protocol_type,
                'file_number': file_number,
                'protocol_number': protocol_number
            })
        
       # Save the data to a jsonl file
        df = pandas.DataFrame(jsonl_data)
        df.to_json(output_path, orient='records', lines=True)

    except Exception as e:
        # Handle any exception
        print(f"An error occurred in main: {e}")
