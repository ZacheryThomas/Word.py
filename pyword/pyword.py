# -*- coding: utf-8 -*-
'''
Main module for pyword
'''

import sys
import time
import argparse
import logging
import os

from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

import docx_handler

class Watcher(object):
    '''
    Class that uses watchdog observer to handle file change events
    '''
    def __init__(self, directory='.'):
        self.directory_to_watch = directory
        self.observer = Observer()

    def run(self):
        event_handler = Handler()
        self.observer.schedule(event_handler, self.directory_to_watch, recursive=True)
        self.observer.start()
        try:
            while True:
                time.sleep(5)
        except:
            self.observer.stop()
            print "Error"

        self.observer.join()


class Handler(FileSystemEventHandler):
    '''
    Class that handles file creation events
    '''
    @staticmethod
    def on_any_event(event):
        if event.event_type == 'created':
            # Take any action here when a file is first created.
            print "Received created event - {}.".format(event.src_path)

            filename = event.src_path

            if filename.endswith(".docx"):
                try:
                    docx_handler.run(filename)
                except Exception as error:
                    print "Error Occured", error


def parse_args(args):
    '''
    Parses command line arguments
    '''
    parser = argparse.ArgumentParser()
    parser.add_argument(
        '--directory',
        action='store',
        dest='directory',
        default='.',
        help='Directory to monitor for docx file changes. (default: current dir script is ran from)'
    )
    return parser.parse_args(args)


def main(args=None):
    '''
    Main entrypoint
    '''
    args = parse_args(args)

    if os.path.isdir(args.directory):
        watcher = Watcher(directory=args.directory)
        watcher.run()
    else:
        logging.error('"%s" is not a directory', args.directory)
        exit(1)


if __name__ == '__main__':
    main(sys.argv[1:])
