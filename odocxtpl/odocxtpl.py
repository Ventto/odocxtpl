#!/usr/bin/env python
#
# MIT License
#
# Copyright (c) 2016 Thomas "Ventto" Venriès
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.
import argparse
import docxtpl
import io
import os
import sys
import yaml

def __getopt():
    parser = argparse.ArgumentParser()
    parser.add_argument('DOCX', help="Input .DOCX file")
    parser.add_argument('YAML', help="Input .YAML file")
    parser.add_argument('OUTPUT', help="OUTPUT .docx file")
    return parser.parse_args()

def main(argv):

    args = __getopt()

    if not os.path.exists(args.docx):
            print(args.docx, ': file not found.\n', file=sys.stderr)
            sys.exit(1)
    if not os.path.exists(args.yaml):
            print(args.yaml, ': file not found.\n', file=sys.stderr)
            sys.exit(1)

    with io.open(args.yaml, 'r') as stream:
        try:
            context = yaml.load(stream)
            if type(context) is not dict:
                print('Error: YAML loaded data are not a dictionary.')
                sys.exit(1)
            doc = docxtpl.DocxTemplate(args.docx)
            doc.render(context)
            doc.save(args.output)
        except yaml.YAMLError as exc:
            print('Error: YAML loaded data are not a dictionary.')
            print(exc)


if __name__ == "__main__":
    main(sys.argv)
