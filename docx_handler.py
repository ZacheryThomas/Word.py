# -*- coding: utf-8 -*-

import subprocess
import pyautogui
import pyperclip
import docx

CTRL_KEY = 'command'
OUTPUT_DELIMITER = "output:"

def select_all_text():
    '''
    Hotkey that selects all on screen text with pyautogui
    '''
    pyautogui.hotkey(CTRL_KEY, 'a')


def delete_all_text():
    '''
    Selects all on screen text and deltes it
    '''
    select_all_text()

    # press delete key
    pyautogui.hotkey('del')


def get_text(filename):
    '''
    Gets all text in word document, returns string
    '''
    doc = docx.Document(filename)
    full_text = []
    for para in doc.paragraphs:
        full_text.append(para.text)
    return '\n'.join(full_text)


def write_output(text):
    '''
    Copies given text to clipboard, then pastes
    '''

    # store output text in clipboard
    pyperclip.copy(text)

    # dump clipboard to wherever cursor is
    pyautogui.hotkey(CTRL_KEY, 'v')


def content_cleaner(text):
    '''
    Cleans content from word fancy characters to python friendly ones
    '''
    replacers = [('“', '"'), ('”', '"'), ('‘', "'"), ('’', "'"), ('-', '-')]
    text = text.encode('utf-8')
    for replacement in replacers:
        text = text.replace(replacement[0], replacement[1])
    return text


def run(filename):
    '''
    Main entrypoint of the module
    '''
    file_content = get_text(filename)

    delimiter_location = file_content.find(OUTPUT_DELIMITER)

    if delimiter_location > -1:
        file_content = file_content[:delimiter_location].rstrip()

    code = content_cleaner(file_content)

    output = ''
    try:
        output = subprocess.check_output(
            ['python', '-c', code],
            stderr=subprocess.STDOUT
            ).decode('utf-8')
    except subprocess.CalledProcessError as exc:
        output = exc.output

    final_output = '''{code}
{delimiter}
{output}
'''.format(code=code, delimiter=OUTPUT_DELIMITER, output=output)

    select_all_text()
    write_output(final_output)
