#!/usr/bin/env python
import argparse
import sys
import email
from pathlib import Path
from email.message import EmailMessage

import aggregator
import fetcher
import status

_ATTACHMENTS_DIRECTORY = Path('output/attachments')
_STATUS_OUTPUT = 'output/status.xlsx'
_ON_DUTY_OUTPUT = 'output/on-duty.xlsx'

def main() -> int:
    parser = argparse.ArgumentParser(description='Divera status tracker')
    subparser = parser.add_subparsers(dest='command')

    parser_fetch = subparser.add_parser('fetch')
    parser_fetch.add_argument('--host', type=str, required=True, help='imap host')
    parser_fetch.add_argument('--email', type=str, required=True, help='Email address')
    parser_fetch.add_argument('--password', type=str, required=True, help='Password')
    parser_fetch.add_argument('--subject', type=str, required=True, help='Subject for filtering')

    parser_aggregator = subparser.add_parser('aggregate')
    parser_aggregator.add_argument('--output', type=str, help='Output path', default=_STATUS_OUTPUT)

    parser_on_duty = subparser.add_parser('on-duty')
    parser_on_duty.add_argument('--input', type=str, help='Input path', default=_STATUS_OUTPUT)
    parser_on_duty.add_argument('--off-duty-keyword', required=True, type=str, help='Keyword to use to filter for off duty')
    parser_on_duty.add_argument('--output', type=str, help='Output path', default=_ON_DUTY_OUTPUT)


    args = parser.parse_args()
    if args.command == 'fetch':
        _ATTACHMENTS_DIRECTORY.mkdir(exist_ok=True, parents=True)
        fetcher.fetch_messages(args.host, args.email, args.password, args.subject, _ATTACHMENTS_DIRECTORY)
        aggregator.aggregate(_ATTACHMENTS_DIRECTORY, _STATUS_OUTPUT)

    elif args.command == 'aggregate':
        if not _ATTACHMENTS_DIRECTORY.exists():
            print('No attachements to aggregate')
            sys.exit(1)
        aggregator.aggregate(_ATTACHMENTS_DIRECTORY, Path(args.output))
    elif args.command == 'on-duty':
        status.calculate_on_duty_percentage(Path(args.input), args.off_duty_keyword, Path(args.output))
    else:
        parser.print_help()
    return 0

if __name__ == '__main__':
    sys.exit(main())
