#!/usr/bin/env python3
# -*- encode: utf8 -*-

from shutil import rmtree
from pathlib import Path
from argparse import ArgumentParser
from oletools.olevba import VBA_Parser, VBA_Project, filter_vba

OFFICE_FILE_EXTENSIONS = (
    '.xlsb', '.xls', '.xlsm', '.xla', '.xlt', '.xlam',  # Excel book with macro
)


def get_args():
    parser = ArgumentParser(description='Extract vba source files from an MS Office file with macro.')
    parser.add_argument('sources', metavar='MS_OFFICE_FILE', type=str, nargs='+',
                        help='Paths to source MS Office file or directory.')
    parser.add_argument('--dest', type=str, default='vba_src',
                        help='Destination directory path to output vba source files [default: ./vba_src].')
    parser.add_argument('--orig-extension', dest='use_orig_extension', action='store_true',
                        help='Use an original extension (.bas, .cls, .frm) for extracted vba source files [default: use .vb].')
    parser.add_argument('--src-encoding', dest='src_encoding', type=str, default='shift_jis',
                        help='Encoding for vba source files in an MS Office file [default: shift_jis].')
    parser.add_argument('--out-encoding', dest='out_encoding', type=str, default='utf8',
                        help='Encoding for generated vba source files [default: utf8].')
    parser.add_argument('--recursive', action='store_true',
                        help='Find sub directories recursively when a directory is specified as the sources parameter.')
    return parser.parse_args()


def get_source_paths(sources, recursive):
    for src in sources:
        p = Path(src)
        if p.is_dir(): # If source is a directory, then find source files under it.
            for file in p.glob("**/*" if recursive else "*"):
                f = Path(file)
                if not f.name.startswith('~$') and f.suffix.lower() in OFFICE_FILE_EXTENSIONS:
                    yield f.absolute()
        else: # If source is a file, then return its absolute path.
            yield p.absolute()


def get_outputpath(parent_dir: Path, filename: str, use_orig_extension: bool):
    extension = filename.split('.')[-1]
    if extension == 'cls':
        subdir = parent_dir.joinpath('class')
    elif extension == 'frm':
        subdir = parent_dir.joinpath('form')
    else:
        subdir = parent_dir.joinpath('module')

    if not subdir.exists():
        subdir.mkdir(parents=True, exist_ok=True)

    return Path(subdir.joinpath(filename + '.vb' if not use_orig_extension else ''))


def extract_macros(parser: VBA_Parser, vba_encoding):

    if parser.ole_file is None:
        for subfile in parser.ole_subfiles:
            for results in extract_macros(subfile, vba_encoding):
                yield results
    else:
        parser.find_vba_projects()
        for (vba_root, project_path, dir_path) in parser.vba_projects:
            project = VBA_Project(parser.ole_file, vba_root, project_path, dir_path, relaxed=False)
            project.codec = vba_encoding
            project.parse_project_stream()

            for code_path, vba_filename, code_data in project.parse_modules():
                yield (vba_filename, code_data)


if __name__ == '__main__':
    args = get_args()

    # Get the root path of destination (if not exists then make it).
    root = Path(args.dest)
    if not root.exists():
        root.mkdir(parents=True)
    elif not root.is_dir():
        raise FileExistsError

    # Get the source MS Office file where extract the vba source files from.
    for source in get_source_paths(args.sources, args.recursive):
        src = Path(source)
        basename = src.stem
        dest = Path(root.joinpath(basename))
        dest.mkdir(parents=True, exist_ok=True)
        rmtree(dest.absolute())
        print('Extract vba files from {source} to {dest}'.format(source=source, dest=dest))

        # Extract vba source files from the MS Office file and save each vba file into the sub directory as of its MS Office file name.
        vba_parser = VBA_Parser(src)
        for vba_filename, vba_code in extract_macros(vba_parser, args.src_encoding):
            vba_file = get_outputpath(dest, vba_filename, args.use_orig_extension)
            vba_file.write_text(filter_vba(vba_code), encoding=args.out_encoding)
            print('[{basename}] {vba_file} is generated.'.format(basename=basename, vba_file=vba_file))
