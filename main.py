# -*- coding: utf-8 -*-

import time

from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

import pyword

class Watcher:
    DIRECTORY_TO_WATCH = "."

    def __init__(self):
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
            print "Received created event - %s." % event.src_path

            filename = event.src_path

            if filename.endswith(".docx"):
                try:
                    pyword.run(filename)
                except Exception as e:
                    print "Error Occured", e


        elif event.event_type == 'modified':
            # Taken any action here when a file is modified.
            print "Received modified event - %s." % event.src_path


if __name__ == '__main__':
    watcher = Watcher()
    watcher.run()