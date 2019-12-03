#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "Enrique_Galindo"

import zipfile
import os
import argparse

def create_parser():
    """
    It creates a parser
    """
    parser = argparse.ArgumentParser(description="searching a given directory for a specific text")
    parser.add_argument("text", help="The text being searched for within each file")
    parser.add_argument("--dir", help="A directory that can be used to search files")
    return parser

def search(name_space):
    search_text = name_space.text
    search_path = name_space.dir
    file_paths = os.listdir(search_path)
    searched_count = 0
    found_count = 0
    print("You are searching {} for dotm files with {}".format(search_path, search_text))
    for f in file_paths:
        searched_count += 1
        full_path = os.path.join(search_path, f)
        """
        this is how we get full paths ^^^
        """
        if zipfile.is_zipfile(full_path):
            with zipfile.ZipFile(full_path) as zip:
                file_names = zip.namelist()
                for name in file_names:
                    if "word/document.xml" in name:
                        with zip.open("word/document.xml", "r") as dotm:
                            for line in dotm:
                                line = line.decode('utf-8')
                                search_result = line.find(search_text)
                                if search_result > -1:
                                    print("match found in {}".format(full_path))
                                    print("...%s..." % line[search_result - 40: search_result + 40])
                                    found_count += 1

    print("Total dotm files searched: %s" % len(file_paths))
    print("Total dotm files matched: %s" % found_count)

def main():
    parser = create_parser()
    name_space = parser.parse_args()
    """
    arg parse creates a namespace of key-value pairs of user input data
    """
    search(name_space)
    
    return 0


if __name__ == '__main__':
    main()
