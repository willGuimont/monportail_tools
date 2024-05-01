import argparse
import os
import pathlib
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
    parser.add_argument('-c', '--courses', nargs='+', required=True, help='Course names, e.g. GLO_4030')
    parser.add_argument('-t', '--teams', nargs='+', type=int, required=True,
                        help='Numbers of teams per course, e.g. 20 34')
    parser.add_argument('-o', '--output', required=False,
                        help='Folder to generate the structure to, defaults to the same folder as `directory`')

    args = parser.parse_args()
    base_comment_file = pathlib.Path(args.base_comment_file).expanduser()
    course_names = args.courses
    teams = args.teams

    assert len(course_names) == len(teams), 'Number of courses and number of teams must be the same'

    if args.output is None:
        args.output = base_comment_file.parent

    output_path = pathlib.Path(args.output).expanduser()

    print(f'Creating {len(course_names)} courses with {sum(teams)} teams in total')

    for course, teams in zip(course_names, teams):
        out_course = output_path.joinpath(course)
        out_course.mkdir(exist_ok=True, parents=True)
        for team in range(1, teams + 1):
            out_file = out_course.joinpath(f'Equipe {team}.docx')
            os.system(f'cp "{str(base_comment_file)}" "{out_file}"')
