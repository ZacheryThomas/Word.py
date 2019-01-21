# -*- coding: utf-8 -*-

import subprocess
import pyautogui
import pyperclip

OUTPUT_DELIMITER = "output:"

def select_all_text():
    '''
    Hotkey that selects all on screen text with pyautogui
    '''
    pyautogui.hotkey('command', 'a')


def get_text():
    '''
    Selects all on screen text and returns what is selected back as a string
    '''
    select_all_text()

    # copy selected text to clipboard and store
    pyautogui.hotkey('command', 'c')
    text = pyperclip.paste()

    return text


def delete_all_text():
    '''
    Selects all on screen text and deltes it
    '''
    select_all_text()

    # press delete key
    pyautogui.hotkey('del')


def write_output(text):
    '''
    Copies given text to clipboard, then pastes
    '''

    # store output text in clipboard
    pyperclip.copy(text)

    # dump clipboard to wherever cursor is
    pyautogui.hotkey('command', 'v')


def content_cleaner(text):
    '''
    Cleans content from word fancy characters to python friendly ones
    '''
    replacers = [('“', '"'), ('”', '"'), ('‘', "'"), ('’', "'"), ('-', '-')]
    text = text.encode('utf-8')
    for replacement in replacers:
        text = text.replace(replacement[0], replacement[1])
    return text


def run():
    '''
    Main entrypoint of the module
    '''
    file_content = get_text()

    output_location = file_content.find(OUTPUT_DELIMITER)

    if output_location > -1:
        file_content = file_content[:output_location].rstrip()

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

    delete_all_text()

    write_output(final_output)
