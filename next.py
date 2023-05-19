import pyttsx3
import docx
import os

def read_file(path_of_file):
    # Load the file
    doc = docx.Document(path_of_file)

    # Extract the text from paragraphs
    text = ""
    for paragraph in doc.paragraphs:
        text += paragraph.text

    return text

def read_aloud(text):
    engine = pyttsx3.init()
    engine.setProperty('rate', 120)
    engine.say(text)

    engine.runAndWait()
    engine.stop()

cwd = os.getcwd()
read = os.listdir(cwd)
for file_to_read in read:
    if file_to_read.endswith('.docx'):
        file_path = os.path.join(cwd, file_to_read)
        print(file_path)
        text = read_file(file_path)
        read_aloud(text)
