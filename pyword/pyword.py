# -*- coding: utf-8 -*-

import sys
import time
import argparse
import logging
import os

from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

import docx_handler

class Watcher:
    DIRECTORY_TO_WATCH = "."

    def __init__(self, directory='.'):
        self.observer = Observer()

    def run(self):
        event_handler = Handler()
        self.observer.schedule(event_handler, self.DIRECTORY_TO_WATCH, recursive=True)
        self.observer.start()
        try:
            while True:
                time.sleep(5)
        except:
            self.observer.stop()
            print "Error"

        self.observer.join()


class Handler(FileSystemEventHandler):

    @staticmethod
    def on_any_event(event):
        if event.event_type == 'created':
            # Take any action here when a file is first created.
            print "Received created event - {}.".format(event.src_path)

            filename = event.src_path

            if filename.endswith(".docx"):
                try:
                    docx_handler.run(filename)
                except Exception as e:
                    print "Error Occured", e


def parse_args(args):
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
    args = parse_args(args)

    if os.path.isdir(args.directory):
        watcher = Watcher(directory=args.directory)
        watcher.run()
    else:
        logging.error('"%s" is not a directory' % args.directory)
        exit(1)



if __name__ == '__main__':
    main(sys.argv[1:])
