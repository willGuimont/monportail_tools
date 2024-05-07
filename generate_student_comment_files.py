import argparse
import os
import pathlib

import pandas as pd

if __name__ == '__main__':
    parser = argparse.ArgumentParser(prog='Generate comment files',
                                     description='Generate comment files for all teams in a course')
    parser.add_argument('base_comment_file',
                        help='Path to the base comment file, e.g. data/Equipe XX.docx')
    parser.add_argument('-c', '--course', required=True, help='Course name, e.g. GLO_4030')
    parser.add_argument('-e', '--excel', required=True,
                        help='Path to the excel results file, e.g. data/Grille des résultats.xlsx')
    parser.add_argument('-o', '--output', required=False,
                        help='Folder to generate the structure to, defaults to the same folder as `directory`')

    args = parser.parse_args()
    base_comment_file = pathlib.Path(args.base_comment_file).expanduser()
    course_name = args.course
    excel_file = pathlib.Path(args.excel).expanduser()

    if args.output is None:
        args.output = base_comment_file.parent

    output_path = pathlib.Path(args.output).expanduser()
    excel_array = pd.read_excel(excel_file, header=11, skipfooter=1, usecols=lambda x: 'Unnamed' not in x)

    out_course = output_path.joinpath(course_name)
    out_course.mkdir(exist_ok=True, parents=True)

    students = excel_array[~excel_array['Nom'].str.contains(r'\(Abandon\)')]

    print(f'Creating files for {len(students)} students in total')
    for i, student in students.iterrows():
        student_name = student['Nom']
        student_id = student['Identifiant']

        out_subfolder = out_course.joinpath(f'{student_name} ({student_id})')
        out_subfolder.mkdir(exist_ok=True, parents=True)

        out_file = out_subfolder.joinpath(f'correction_{student_id}.docx')
        os.system(f'cp "{str(base_comment_file)}" "{out_file}"')
