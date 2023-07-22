from docx2python import docx2python
import os

# -- Setting parsing parameters --
answer_identifier = "Ans"
question_number = 1
options = ["A", "B", "C", "D", "E"]

# -- Specifying a directory for source files --
directory_name = f'source-files'
directory_path = f'{os.getcwd()}/{directory_name}'
# print(f'Directory: {directory_path}')

# -- Check if selected path is a file or a directory --
# if os.path.isdir(directory_path):  
#     print("It is a directory")  
# elif os.path.isfile(file_name):  
#     print("It is a normal file")  
# else:  
#     print("It is a special file (socket, FIFO, device file)" )

# -- Specifying a file for parsing --
# file_name = f'source-files/20 Chapter Pharmacokinetics Q&A ver 1.docx'

# file_name = f'{directory_path}/{os.listdir(directory_path)[0]}'


# -- Methods --
def replace(line, find, replace):
    '''
    This function takes a line of text and replaces a keyword with another.
    
    :param line: This is the line of text to search in
    :param find: This is the keyword to look for in the line of text
    :param replace: This is the keyword to be replaced with

    '''
    return line.replace(find, replace)

# document = docx2python(file_name)

final_text = []

final_document = ""

element_list = []

if __name__ == '__main__':
    for file in os.listdir(directory_path):
        file_name = os.fsdecode(file)
        print(f'File name: {file_name}')
        if file_name.endswith(".docx"):
            try: 
                document = docx2python(f'{directory_path}/{file_name}')
                final_text = []

                final_document = ""
                element_list = []

                for element in document.body:
                    element_list.append(element)
                    
                text_content = []
                
                text_content2 = [item[0] for sublist in element_list for item in sublist]

                text_content = [item for sublist in text_content2 for item in sublist]
                
                missing_answers = 0
                tips_counter = 0

                for index, line in enumerate(text_content):
                    if line.lstrip().startswith(f'{question_number}.'):
                        question_number += 1
                        # -- Print the question without the question number --
                        # print(level.lstrip().removeprefix(f'{question_number}.').lstrip())
                        
                    if line.lstrip().startswith(f'Ans'):
                        answer = line.lstrip().removeprefix(f'Ans.').removeprefix(f'Ans:').lstrip()
                        if not any([option in answer for option in options]):
                            missing_answers+= 1
                        # -- Print the answer option --
                        # print(level.lstrip().removeprefix(f'Ans.').removeprefix(f'Ans:').lstrip())

                    if "tips" in line.lstrip().lower():
                        tips_counter+= 1
                        line = line.lstrip().replace("Tips:", "Exp: Tips:")

                    if line.lstrip().startswith(f'----media/'):
                        continue

                    final_document += line.lstrip() + "\n"

                output_path = f'output/{file_name}.txt'
                if missing_answers > 0:
                    print(f'Missing Ans: {missing_answers}')
                    with open(output_path, 'w') as f:
                        f.write(f'Missing Ans: {missing_answers}')
            except Exception as e:
                output_path = f'output/{file_name}.txt'
                with open(output_path, 'w') as f:
                    f.write(f'Exception:\n {e}')
            else:
                # print(final_document)
                with open(output_path, 'w') as f:
                    f.write(final_document)
            continue
        elif file_name.endswith(".doc"):
            output_path = f'output/{file_name}.txt'
            with open(output_path, 'w') as f:
                    f.write(f'Only .docx documents can be parsed')
        else:
            continue
