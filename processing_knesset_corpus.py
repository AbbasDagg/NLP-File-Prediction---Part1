from docx import Document
import zipfile
import os
import pandas
import re
import json
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

def get_all_docx_in_current_foleder(folder_path):
    try:

        info =[]
        # go throw all the documants in the directory
        current_path =os.path.join(os.getcwd(),folder_path)
        
        for file_number,filename in enumerate(os.listdir(current_path)):
            if filename.endswith('docx'):
                
                attributes = filename.split('_')# get the attributes of the file
                #then assigen it to the right variable
                if attributes[1] == 'ptv':  # committee or plenary
                        type = 'committee'
                elif attributes[1] == 'ptm':
                        type = 'plenary'
                else:
                    type = '-1'

                text = Document(os.path.join(current_path,filename))
                
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
        print(f'Exception in get_all_docx_in_current_foleder: {e}')

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

def is_underlined(par):
    try:
        # Check if all text runs are underlined
        all_underlined = True  # Assume all are underlined unless proven otherwise
        for run in par.runs:
            if not run.underline:
                all_underlined = False  # If any run is not underlined, then it's not fully underlined

        if all_underlined:
            return True

        # If any run is underlined, consider it partially underlined
        for run in par.runs:
            if run.underline:
                return True

        # Check the paragraph style for underline attributes
        current_style = par.style
        while current_style:
            if current_style.font.underline:
                return True  # If underline is defined at the style level

            # Check base styles for inherited underline
            current_style = current_style.base_style

        return False  # Default to not underlined
    except Exception as e:
        print(f"Exception in is_underlined: {e}")
        return False

def clear_name(name): # clean
    try:
        name = name.strip()
        comps = name.split(' ')
        new_name = ""
        open_parentheses =False 
        for comp in comps: # go throw all the components of the name
            if comp == '':
                continue

            if '(' in comp:
                open_parentheses = False
                continue
            if open_parentheses:
                continue
            if '"' not in comp and '”' not in comp:# this means that the component is not a short cut
                if '\'' == comp[-1] and len(comp) <4: #then this means that it might be the person numbering of his possition
                    new_name = '' # is so then just throw all of what was before it
                    continue
                    #if the number is more than 2 digits in hebrew this means this code will not capture it
                    #another problem if someone name has 2 latters and ends with ' then this code will not capture it 
                    #if position came after the name then this code will not capture it

                if ")" in comp: #if we have open_parentheses then we dont take it
                    if "(" in comp:# if we have closing parentheses then we dont take it
                        continue
                    else:# else we must take the following comps until we see closing parentheses
                        open_parentheses = True
                elif comp == "-" or comp == '–' or comp == '~' or comp ==',':#if the name has a dash then take the first part
                    break
                else:
                    new_name += comp+" "

        new_name = new_name.strip()
        if ',' in new_name:
            return ''
        if new_name != "" and new_name.find(':')+1 == len(new_name):
            new_name = new_name[:-1]

        return new_name.strip()
    except Exception as e:
        print(f'Exception in clear_name: {e}')

def split_paragrph(par):
    try:
        new_text = ''
        sentence_seprator = '.؟!:;' #we split according to thoughs
        we_are_in_quets = False
        par_parts = par.text.strip().split(' ')
        sentece_list = []
        for part in par_parts:
            
            if part =='':# if the text is empty there is nothing to do 
                continue
            new_text += part +" "
            if '"' == part[0] and we_are_in_quets == False: 
                we_are_in_quets = True
            if '"' == part[-1] or (len(part)>=2 and part[-2] == '"' and part[-1] in sentence_seprator):
                # if the last charerchter is a quation mark or the second to last charechter and it has a seprator then this means that this is the end of the quation
                we_are_in_quets = False

            if part[-1]  in sentence_seprator or (len(part)>=2 and part[-2] in sentence_seprator):
                # if the sentence ends with are the sentence_seprator and then '\"' we catch that 
                if we_are_in_quets == False:
                    # we just add if the sentence_seprator is not in the quets
                    sentece_list.append(new_text.strip())
                    new_text = ''
        #if we start a quet but didnt end it then save the text
        if we_are_in_quets:
            sentece_list.append(new_text)
        return sentece_list
    except Exception as e:
        print(f'exception in split_paragrph: {e}')

if __name__ == "__main__":
    try:
        print(sys.argv)
        if len(sys.argv) !=3:
            print('Incorrect input, please enter the folder path and the output path.')
            sys.exit(1)

        # Add Check if folder is valid

       # CHECK IF ITS IN CORRECT JSONL FORMAT
        folder_path = sys.argv[1]
        output_path = sys.argv[2]
        #info = get_docx(folder_path)
        info = get_all_docx_in_current_foleder(folder_path)
        cpy = info.copy()
        print(info[0])
        target_words = ["הישיבה ה", "פרוטוקול מס'"]  # Search for the protocal number
        jsonl_data = []
        names = []

        for doc_num, doc in enumerate(info):
            knesset_number = doc['knesset number']
            protocol_type = doc['type']
            file_number = doc['file_number']
            protocol_number = -1 # Default value
    
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
            
            speaker_name ='' #name of the current_speaker
            speaker_text = {} # a dictionary of text (speaker ,all his text)
            # Extract speakers and text
            for par in doc['text'].paragraphs:
                text = par.text.strip() # Remove leading and trailing spaces
                if text.startswith('<') or text.startswith('>'): # Sometimes the text starts with < and ends with >, its probably caused by the conversion from doc to docx so we remove it
                    text = text[1:-1] 

            #if doc['file_number'] == '302840':
            #for i in range(100):
            names.append({'docx':doc['file_number']})
            for par in doc['text'].paragraphs:
                text = par.text
                #text = doc['text'].paragraphs[i].text
                if text.startswith('<') or text.startswith('>'): # Sometimes the text starts with < and ends with >, its probably caused by the conversion from doc to docx so we remove it
                    text = text[1:-1] 
                
                index = text.strip().find(":")
                if index>=0:  # if the last char is : and the whole text is underlined then this is a speaker
                    if index== len(text) -1 and is_underlined(par):
                        #print(f"Name: {text}")
                        #print('---------yes')
                        #continue
                        clear_nam = clear_name(text)
                        names.append({'names':text,
                                      'clear_name':clear_nam,})
                #print(text)
                #print('---------no')

            # Put this in the correct place
            # Append the data to the jsonl_data list

            jsonl_data.append({
                'knesset_number': knesset_number,
                'protocol_type': protocol_type,
                'file_number': file_number,
                'protocol_number': protocol_number
            })
        
        with open(output_path, 'w', encoding='utf-8') as jsonl_file:
            for data_item in names: # change back
                # Convert the dictionary to a JSON-formatted string
                json_line = json.dumps(data_item, ensure_ascii=False)
        
                # Write the JSON string to the file with a newline separator
                jsonl_file.write(json_line + '\n')

    except Exception as e:
        # Handle any exception
        print(f"An error occurred in main: {e}")
