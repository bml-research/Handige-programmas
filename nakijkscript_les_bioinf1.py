#!/usr/bin/env python3
"""
Name: nakijkscript.py
Purpose: nakijken bioinfo opgaven.
Author: Jurre Hageman and Mark Sibbald
Created: 2015-03-21
Updated: 2020-01-27
"""

#imports
import xlsxwriter
import os.path
import argparse
import time
import csv
import sys
from bs4 import BeautifulSoup
import warnings
warnings.filterwarnings('ignore', category=
                        UserWarning, module='bs4')
t0 = time.time()


def args():
    "parses command line arguments"
    parser = argparse.ArgumentParser(description="checks student answers")
    parser.add_argument("in_file", help="the path to the File with the input")
    parser.add_argument("out_file", help="the path to the File with the output")
    parser.add_argument("antwoord_file", help="the answer file")
    args = parser.parse_args()
    return args


def check_anwers(student_answers, model):
    score = 0
    score_dict = {}

    for num, answer in enumerate(student_answers):
        answer = answer.strip()
        if answer in model[num+1]:
            score += 1
            score_dict.update({num+1:1})
        else:
            score_dict.update({num+1:student_answers[num]})
    res = { 'score' : score, 'score_per_question' : score_dict}
    return res


def delete_html_tags(file_name):
    f = open(file_name, 'rt', encoding="utf16") # opens the csv file
    reader = csv.reader(f, delimiter='\t', quoting=csv.QUOTE_NONE)  # creates the reader object
    data = []

    for line in reader:
        new_line = []
        if line[0] != "Username" and line[0] != "Gebruikersnaam":
            for element in line:
                soup = BeautifulSoup(element, 'html.parser')
                strip = soup.get_text()
                strip = strip.replace('"','')
                strip = strip.strip() # remove quotes and whitespace
                new_line.append(strip)
            data.append(new_line)
    return data


def read_answers(antwoordFile):
    antwoorden = {}
    for num, line in enumerate(open(antwoordFile, "r")):
        line = line.rstrip("\n").split(";")
        item = [i.strip() for i in line]
        antwoorden.update({num+1:item})
    return antwoorden


def create_student_list(data):
    students = []
    for line in data:
        if line[0] != "Username" and line[0] != "Gebruikersnaam":
            naw_number = line[0]
            sur_name = line[1]
            first_name = line[2]

            # next is unique for lessons bioinf1; could be used for the exams to sort on class A and B
            if line[-4] == "":
                group_inf = ['Niet', 'ingevuld']
            else:
                group_inf = line[-4].split(" ")
                group_inf = [i.strip().strip('"') for i in group_inf]
            group = group_inf[0]
            subgroup = group_inf[1]

            answers = line[5::6]
            answers = answers[:-1]
            # modify number 1
            for num, answer in enumerate(answers):
                if answer.strip() == "1":
                    answers[num] = "number: 1"            
            students.append({
                'naw_number': naw_number,
                'sur_name' : sur_name,
                'first_name' : first_name,
                'group' : group,
                'subgroup' : subgroup,
                'answers' : answers})
    return students



def write_xlsx(students, outFile):
    try:
        ofile = open(outFile, "w")
        workbook = xlsxwriter.Workbook(outFile)
        worksheet = workbook.add_worksheet('Nakijken Bioinformatica 1')

        info = ["Studentnummer", "Achternaam", "Voornaam", "Klas", "Groep", "Aantal vragen", "Score"]
        vraag = ["vraag " + str(i+1) for i in range(len(students[0]['answers']))]
        headers = info + vraag
        results_list = []
        for student in students:
            data = []
            data.append(student['naw_number'])
            data.append(student['sur_name'])
            data.append(student['first_name'])
            data.append(student['group'])
            data.append(student['subgroup'])
            data.append(int(len(student['answers'])/2))
            data.append(student['student_score']['score'])
            for num, antwoord in sorted(student['student_score']['score_per_question'].items()):
                data.append(antwoord)
            results_list.append(data)

        results_list.insert(0, headers)
        for row_num, row_data in enumerate(results_list):
            for col_num, col_data in enumerate(row_data):
                worksheet.write(row_num, col_num, col_data)
        workbook.close()
    finally:
        ofile.close()



def main():
    comm_args = args()
    in_file = comm_args.in_file
    out_file = comm_args.out_file
    antwoord_file = comm_args.antwoord_file
    text_stripped = delete_html_tags(in_file)
    students = create_student_list(text_stripped)
    answer_model = read_answers(antwoord_file)

    for student in students:
        print("processing: {}, {}, {}".format(student['naw_number'], student['sur_name'], student['first_name']))
        if len(answer_model) != len(student['answers']):
            print("WAARSCHUWING: aantal antwoorden antwoordmodel: ({}) is niet gelijk aan het aantal antwoorden ({}) van student: {} {}".format(len(answer_model),len(student['answers']), student['sur_name'], student['first_name']))
            sys.exit(1)
        student_score = check_anwers(student['answers'], answer_model)
        student['student_score'] = student_score
    if os.path.isfile("{}".format(out_file)):
        print("File {} bestaat al".format(out_file))
        #response = 'y'
        response = input("Overschrijven? y/n: ")
        if response == "y":
            write_xlsx(students, out_file)
        else:
            sys.exit(0)
    else:
        write_xlsx(students, out_file)
    print('\nRuntime lasted {0:.1f} sec'.format(time.time()-t0))
        
if __name__ == "__main__":
    main()
