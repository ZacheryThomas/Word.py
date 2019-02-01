# -*- coding: utf-8 -*-
'''
Main module for pyword
'''

import sys
import time
import argparse
import subprocess
import logging
import os

from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

import docx_handler
from util import pyword_logger


class Watcher(object):
    '''
    Class that uses watchdog observer to handle file change events
    '''
    def __init__(self, path='.'):
        self.path_to_watch = path
        self.observer = Observer()

    def run(self):
        event_handler = Handler()
        self.observer.schedule(event_handler, self.path_to_watch, recursive=True)
        self.observer.start()
        try:
            while True:
                time.sleep(5)
        except Exception as exception:
            self.observer.stop()
            pyword_logger.error("Exception thrown: %s", exception)

        self.observer.join()


class Handler(FileSystemEventHandler):
    '''
    Class that handles file creation events
    '''
    @staticmethod
    def on_any_event(event):
        if event.event_type == 'created':
            # Take any action here when a file is first created.
            pyword_logger.info("Received created event - %s.", event.src_path)

            filename = event.src_path

            if filename.endswith(".docx"):
                try:
                    docx_handler.run(filename)
                except Exception as exception:
                    pyword_logger.error("Exception thrown: %s", exception)


def parse_args(args):
    '''
    Parses command line arguments
    '''
    parser = argparse.ArgumentParser(description='Run python in word')
    parser.add_argument('--verbose', action='store_true', dest='verbose',
                        help='Increases verbosity of output')
    parser.add_argument(
        '--path',
        action='store',
        dest='path',
        default='.',
        help='Directory or file to monitor for docx file changes. (default: current dir script is ran from)'
    )
    return parser.parse_args(args)


def main(args=None):
    '''
    Main entrypoint
    '''
    pyword_logger.info('Running pyword!')
    args = parse_args(args)

    if args.verbose:
        pyword_logger.setLevel(logging.DEBUG)

    pyword_logger.debug('args: %s', args)

    path = args.path
    if (os.path.isdir(path)) or (os.path.isfile(path)):
        is_file = os.path.isfile(path)
        if is_file:
            if path.endswith('.docx'):
                pyword_logger.info('''"%s" is a file. I'll open it for you!''', path)
                subprocess.call(['open', path])
            else:
                pyword_logger.error('"%s" is a file, but does not end in .docx', path)

        path = os.path.dirname(os.path.abspath(path))
        pyword_logger.info('Starting pyword monitoring on "%s"', path)
        watcher = Watcher(path=path)
        watcher.run()
    else:
        pyword_logger.error('"%s" is not a directory or a file', path)
        exit(1)
    pyword_logger.info('Bye!')


if __name__ == '__main__':
    main(sys.argv[1:])
