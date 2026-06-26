from __future__ import annotations

import argparse
from pathlib import Path


def _cmd_app(args: argparse.Namespace) -> int:
    from .app import main as app_main

    argv: list[str] = []
    if args.document is not None:
        argv.append(str(args.document))
    return app_main(argv)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="python3 -m prose")
    subparsers = parser.add_subparsers(dest="command", required=True)

    app_parser = subparsers.add_parser("app", help="launch the GTK/Libadwaita Prose app")
    app_parser.add_argument("document", nargs="?", type=Path, help="optional Writer .odt document to open")
    app_parser.set_defaults(func=_cmd_app)
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    return int(args.func(args))
