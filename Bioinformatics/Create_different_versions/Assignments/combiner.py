#!/usr/bin/env python3
"""
Name: combine_word.py
Purpose: combine word files.
Author: Jurre Hagema2
Updated: 2023-01-12
"""

#imports
import argparse
import os
from docxcompose.composer import Composer
from docx import Document
import itertools


def args():
    "parses command line arguments"
    parser = argparse.ArgumentParser(description="checks student answers")
    parser.add_argument("in_file", help="the path to the folder with the input")
    parser.add_argument("out_file", help="the path to the folder with the output")
    args = parser.parse_args()
    return args


def iterate_folders(root_folder):
    file_paths = []
    for subdir, dirs, files in os.walk(root_folder):
        if not subdir == "input":
            file_paths.append([subdir, dirs, files])
    return sorted(file_paths)

def create_combos(files_to_process):
    used = []
    total = []
    for i in files_to_process:
        exercise_path = i[0]
        if not exercise_path in used:
            new_paths = []
            for j in i[-1]:
                path = os.path.join(exercise_path, j)
                new_paths.append(path)
            used.append(exercise_path)
            total.append(new_paths)
    combinations = [p for p in itertools.product(*total)]
    return combinations

def create_dir(path):
    isExist = os.path.exists(path)
    if not isExist:
        os.makedirs(path)


def combine_files(combinations, source, destination):
    for num, item in enumerate(combinations):
        master = Document(os.path.join(source, "voorblad.docx"))
        composer = Composer(master)
        for doc in item:
            new_doc = Document(doc)
            composer.append(new_doc)
        doc_path = os.path.join(destination, "version{}.docx".format(num + 1))
        print("Working on {}".format(doc_path))
        composer.save(doc_path)
        


def main():
    comm_args = args()
    in_path = comm_args.in_file
    out_path = comm_args.out_file
    create_dir(out_path)
    files_to_process = iterate_folders(in_path)
    combinations = create_combos(files_to_process)
    combine_files(combinations, in_path, out_path)
    print("done")


if __name__ == "__main__":
    main()