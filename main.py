#!/usr/bin/env python
import argparse
import sys
import email
from pathlib import Path
from email.message import EmailMessage

import aggregator
import fetcher

_ATTACHMENTS_DIRECTORY = Path('attachments')
_AGGREGATED_OUTPUT_PATH = Path('status.xlsx')

def main() -> int:
    parser = argparse.ArgumentParser(description='Divera status tracker')
    subparser = parser.add_subparsers(dest='command')

    parser_fetch = subparser.add_parser('fetch')
    parser_fetch.add_argument('--host', type=str, required=True, help='imap host')
    parser_fetch.add_argument('--email', type=str, required=True, help='Email address')
    parser_fetch.add_argument('--password', type=str, required=True, help='Password')
    parser_fetch.add_argument('--subject', type=str, required=True, help='Subject for filtering')

    parser_aggregator = subparser.add_parser('aggregate')
    parser_aggregator.add_argument('--output', type=str, help='Output path')


    args = parser.parse_args()
    if args.command == 'fetch':
        _ATTACHMENTS_DIRECTORY.mkdir(exist_ok=True)
        messages = fetcher.fetch_messages(args.host,  args.email, args.password, args.subject)
        fetcher.dettach_files_from_messages(messages, _ATTACHMENTS_DIRECTORY)
    elif args.command == 'aggregate':
        if not _ATTACHMENTS_DIRECTORY.exists:
            print('No attachements to aggregate')
            sys.exit(1)
        output_path = Path(args.output) if args.output else _AGGREGATED_OUTPUT_PATH
        aggregator.aggregate(_ATTACHMENTS_DIRECTORY, output_path)
    else:
        parser.print_help()
    return 0

if __name__ == '__main__':
    sys.exit(main())  # next section explains the use of sys.exit
