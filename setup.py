# -*- coding: utf-8 -*-
import os

from setuptools import find_packages
from setuptools import setup

setup(
    name='pyword',
    version='0.0.1',
    description='Run python in word',
    setup_requires='setuptools',
    entry_points={
        'console_scripts': ['pyword=pyword.pyword:main']},
    packages=find_packages(),
    install_requires=[
        'PyAutoGUI>=0.9.41',
        'pyobjc>=5.1.2',
        'pyobjc-core>=5.1.2',
        'pyperclip>=1.7.0',
        'watchdog>=0.9.0',
        'python-docx>=0.8.10'
    ]
)