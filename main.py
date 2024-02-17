#!/usr/bin/env python
import argparse
import fetcher
import sys
import email
from pathlib import Path
from email.message import EmailMessage

DETTACHED_FILES_DIRECTORY = Path('dettached_files')
        
def main() -> int:
    parser = argparse.ArgumentParser(description='Divera status tracker')
    subparser = parser.add_subparsers(dest='command')

    parser_fetch = subparser.add_parser('fetch')
    parser_fetch.add_argument('--host', type=str, required=True)
    parser_fetch.add_argument('--email', type=str, required=True)
    parser_fetch.add_argument('--password', type=str, required=True)
    parser_fetch.add_argument('--subject', type=str, required=True)

    args = parser.parse_args()
    if args.command == 'fetch':
        DETTACHED_FILES_DIRECTORY.mkdir(exist_ok=True)
        messages = fetcher.fetch_messages(args.host,  args.email, args.password, args.subject)
        fetcher.dettach_files_from_messages(messages, DETTACHED_FILES_DIRECTORY)
    else:
        parser.print_help()
    return 0

if __name__ == '__main__':
    sys.exit(main())  # next section explains the use of sys.exit
