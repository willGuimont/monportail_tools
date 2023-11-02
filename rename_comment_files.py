import argparse
import os
import pathlib
from typing import Optional


def parse_team_num(filepath: pathlib.Path) -> Optional[int]:
    try:
        return int(filepath.stem.split('_')[0])
    except ValueError or IndexError:
        try:
            return int(filepath.stem.split(' ')[1])
        except ValueError or IndexError:
            return None


if __name__ == '__main__':
    parser = argparse.ArgumentParser(prog='monPortail correction tool',
                                     description='Creates the directory structure required to upload correction files to students via monPortail')
    parser.add_argument('directory',
                        help='Directory containing all comment files, each file must start with [team_number]_')
    parser.add_argument('-o', '--output', required=False,
                        help='Folder to generate the structure to, defaults to the same folder as `directory`')

    args = parser.parse_args()
    if args.output is None:
        args.output = args.directory

    comment_path = pathlib.Path(args.directory).expanduser()
    output_path = pathlib.Path(args.output).expanduser()

    for file in comment_path.iterdir():
        team_num = parse_team_num(file)
        if team_num is not None:
            team_path = output_path.joinpath(f'Equipe {team_num}')
            team_path.mkdir(parents=True, exist_ok=True)
            os.system(f'cp "{str(file)}" "{team_path}/"')
