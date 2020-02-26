#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "Avelyno Koumaka"


import zipfile
import argparse
import os
import sys


def parserCreator():
    parser = argparse.ArgumentParser()
    parser.add_argument('--dir', action="store", default=os.getcwd())
    parser.add_argument('text')
    return parser


def main():
    parser = parserCreator()
    args = parser.parse_args()

    if not args:
        parser.print_usage()
        sys.exit(1)

    dir = args.dir
    string = args.text
    searchedFiles = 0
    matchedFiles = 0

    if os.getcwd() == dir:
        print("Searching directory {} for text '{}' ...".format(dir, string))
        for files in os.listdir(os.getcwd()):
            if os.path.isfile(os.path.join(os.getcwd(), files)) and not files.startswith("."):
                searchedFiles += 1
                file_path = os.path.join(dir, files)
                content = ''
                with open(files, 'r') as read_file:
                    for string in read_file.readlines():
                        content += string.replace('\n', '\\n')
                    i = content.find(string)
                    if i > 0:
                        matchedFiles += 1
                        print('Match found in file {}'.format(file_path))
                        if i < 40:
                            print('   ...{}{}...'.format(
                                content[:i], content[i:i+len(string)+39]))
                            continue
                        print('   ...{}{}...'.format(
                            content[i-40:i], content[i:i+len(string)+39]))
        print('# of searched files : {}'.format(searchedFiles))
        print('# of matched files : {}'.format(matchedFiles))

    elif os.path.isdir(dir):
        print("Searching directory {} for text '{}' ...".format(dir, string))
        [(dirpath, dirnames, filenames)] = list(os.walk(dir))
        for dotm in filenames:
            if dotm.endswith('.dotm'):
                searchedFiles += 1
                file_path = os.path.join(dir, dotm)
                with zipfile.ZipFile(file_path) as zf:
                    segment = zf.read('word/document.xml')
                    i = segment.find(string)
                if i > 0:
                    matchedFiles += 1
                    print('Match found in file {}'.format(file_path))
                    print('   ...{}{}...'.format(
                        segment[i-40:i], segment[i:i+len(string)+39]))
        print('# of searched files: {}'.format(searchedFiles))
        print('# of matched files: {}'.format(matchedFiles))
    else:
        print('{} is not a valid directory'.format(dir))


if __name__ == '__main__':
    main()
