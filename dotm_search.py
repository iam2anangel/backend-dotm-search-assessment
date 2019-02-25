#!/usr/bin/env python #has to be first line in python script a hint to your macos/linux meaning all the following code is going to be python script for python interpreterls

# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "iam2anangel"  # Jen Browning

import os
import sys
import zipfile
import argparse
import re


def main(directory, text_to_search):
    files_searched = 0
    files_matched = 0
    cwd = os.getcwd()
    file_list = os.listdir(directory)
    print("Searching directory {} for text '{}' ...".format(
        directory, text_to_search))
    for files in file_list:
        files_searched += 1
        full_path = os.path.join(directory, files)
        if not files.endswith(".dotm"):
            print("this isn't a dotm {}".format(files))
            continue
        # checks to see if the file is a zipfile
        if not zipfile.is_zipfile(full_path):
            print("this isn't a zipfile {}".format(full_path))
            continue
        with zipfile.ZipFile(full_path, "r") as zipped:
            # opens, closes, and reads the zipfile
            toc = zipped.namelist()
            if "word/document.xml" in toc:
                with zipped.open("word/document.xml", "r") as doc:
                    for line in doc:
                        # if you get a -1 it's not there
                        i = line.find(text_to_search)
                        if i >= 0:
                            files_matched += 1
                            print("...{}...".format(line[i - 40:i + 40]))
    print("Files matched: {}".format(files_matched))
    print("Files searched: {}".format(files_searched))


def create_parser():
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "directory", help="The directory of dotm files to search")
    parser.add_argument("text_to_search", help="The text to search for")
    return parser


if __name__ == '__main__':
    parser = create_parser()
    namespace = parser.parse_args()
    print(namespace)
    main(namespace.directory, namespace.text_to_search)
