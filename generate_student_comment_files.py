import argparse
import os
import pathlib
import pandas as pd
from typing import Optional


def parse_team_num(filepath: pathlib.Path) -> Optional[int]:
    try:
        return int(filepath.stem.split('_')[0])
    except ValueError or IndexError:
        return None


if __name__ == '__main__':
    parser = argparse.ArgumentParser(prog='Generate comment files',
                                     description='Generate comment files for all teams in a course')
    parser.add_argument('base_comment_file', help='Path to the base comment file, e.g. data/Equipe XX.docx')
    parser.add_argument('-c', '--course', required=True, help='Course name, e.g. GLO_4030')
    parser.add_argument('-e', '--excel', required=True, help='Path to the excel results file, e.g. data/Grille des résultats.xls')
    parser.add_argument('-o', '--output', required=False,
                        help='Folder to generate the structure to, defaults to the same folder as `directory`')

    args = parser.parse_args()
    base_comment_file = pathlib.Path(args.base_comment_file).expanduser()
    course_name = args.course
    excel_file = pathlib.Path(args.excel).expanduser()

    assert args.excel.endswith('.csv'), "As a current workaround, please convert the excel file to a .csv"

    if args.output is None:
        args.output = base_comment_file.parent

    output_path = pathlib.Path(args.output).expanduser()

    with open(excel_file, 'r', encoding='utf8') as f_excel:
        excel_array = pd.read_csv(f_excel, header=11, skipfooter=1, engine='python')

    out_course = output_path.joinpath(course_name)
    out_course.mkdir(exist_ok=True, parents=True)

    students = [list(row[1:3]) for idx, row in excel_array.iterrows()]

    print(f'Creating files for {len(students)} students in total')

    for student in students:
        out_subfolder = out_course.joinpath(f'{student[0]} ({student[1]})')
        out_subfolder.mkdir(exist_ok=True, parents=True)

        out_file = out_subfolder.joinpath(f'correction_{student[1]}.docx')
        os.system(f'cp "{str(base_comment_file)}" "{out_file}"')
