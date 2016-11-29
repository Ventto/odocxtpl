#!/usr/bin/env python

import os.path
import io
import sys
import yaml
from docxtpl import DocxTemplate

def get_status(filepath):
    if not os.path.exists(filepath):
        print('Error: File not found.\n')
        sys.exit(1)
        return filepath

def main(argv):
    argc = len(argv) - 1
    if argc < 3 or argc > 3:
        print('Usage: docxtemplate [FILE] [YAML] [OUTPUT] \n')
        sys.exit(0)

    for i in range(1, argc):
        if not os.path.exists(argv[i]):
            print(argv[i], ': file not found.\n', file=sys.stderr)
            sys.exit(1)

    with io.open(argv[2], 'r') as stream:
        try:
            context = yaml.load(stream)
            if type(context) is not dict:
                print('Error: YAML loaded data are not a dictionary.')
                sys.exit(1)
            doc = DocxTemplate(argv[1])
            doc.render(context)
            doc.save(argv[3])
        except yaml.YAMLError as exc:
            print('Error: YAML loaded data are not a dictionary.')
            print(exc)

if __name__ == "__main__":
    main(sys.argv)
