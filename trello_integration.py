#!/usr/bin/env python3
"""Command-line Trello API integration.

The Trello API uses an app key plus a member token for normal REST requests.
Run `auth-url` first, authorize the app in your browser, then paste the token
into `.env` as TRELLO_TOKEN.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Any


API_BASE = "https://api.trello.com/1"
AUTH_BASE = "https://trello.com/1/authorize"
ENV_PATH = Path(__file__).with_name(".env")


def load_dotenv(path: Path = ENV_PATH) -> None:
    if not path.exists():
        return

    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue

        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        os.environ.setdefault(key, value)


def require_env(name: str) -> str:
    value = os.environ.get(name, "").strip()
    if not value:
        raise SystemExit(f"Missing {name}. Add it to {ENV_PATH} or set it in your shell.")
    return value


def build_auth_url(args: argparse.Namespace) -> str:
    key = require_env("TRELLO_API_KEY")
    params = {
        "expiration": args.expiration,
        "name": args.app_name,
        "scope": args.scope,
        "response_type": "token",
        "key": key,
    }
    return f"{AUTH_BASE}?{urllib.parse.urlencode(params)}"


def trello_request(
    method: str,
    path: str,
    *,
    params: dict[str, Any] | None = None,
    data: dict[str, Any] | None = None,
) -> Any:
    key = require_env("TRELLO_API_KEY")
    token = require_env("TRELLO_TOKEN")
    params = dict(params or {})
    params.update({"key": key, "token": token})

    url = f"{API_BASE}{path}?{urllib.parse.urlencode(params)}"
    encoded_data = None
    if data is not None:
        encoded_data = urllib.parse.urlencode(data).encode("utf-8")

    request = urllib.request.Request(url, data=encoded_data, method=method)
    try:
        with urllib.request.urlopen(request, timeout=30) as response:
            body = response.read().decode("utf-8")
    except urllib.error.HTTPError as exc:
        error_body = exc.read().decode("utf-8", errors="replace")
        raise SystemExit(f"Trello API error {exc.code}: {error_body}") from exc
    except urllib.error.URLError as exc:
        raise SystemExit(f"Could not reach Trello API: {exc.reason}") from exc

    return json.loads(body) if body else None


def print_json(value: Any) -> None:
    print(json.dumps(value, indent=2, sort_keys=True))


def list_boards(_: argparse.Namespace) -> None:
    boards = trello_request(
        "GET",
        "/members/me/boards",
        params={"fields": "name,url,closed", "filter": "open"},
    )
    print_json(boards)


def list_lists(args: argparse.Namespace) -> None:
    lists = trello_request(
        "GET",
        f"/boards/{args.board_id}/lists",
        params={"fields": "name,closed", "filter": "open"},
    )
    print_json(lists)


def list_cards(args: argparse.Namespace) -> None:
    cards = trello_request(
        "GET",
        f"/lists/{args.list_id}/cards",
        params={"fields": "name,desc,url,due,closed", "filter": "open"},
    )
    print_json(cards)


def create_card(args: argparse.Namespace) -> None:
    payload = {
        "idList": args.list_id,
        "name": args.name,
        "desc": args.description or "",
    }
    if args.due:
        payload["due"] = args.due

    card = trello_request("POST", "/cards", data=payload)
    print_json(card)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Trello API integration CLI")
    subparsers = parser.add_subparsers(dest="command", required=True)

    auth = subparsers.add_parser("auth-url", help="Print a Trello token authorization URL")
    auth.add_argument("--app-name", default="trello-ppt")
    auth.add_argument("--scope", default="read,write")
    auth.add_argument("--expiration", default="never")
    auth.set_defaults(func=lambda args: print(build_auth_url(args)))

    boards = subparsers.add_parser("boards", help="List your open Trello boards")
    boards.set_defaults(func=list_boards)

    lists = subparsers.add_parser("lists", help="List open lists on a board")
    lists.add_argument("board_id")
    lists.set_defaults(func=list_lists)

    cards = subparsers.add_parser("cards", help="List open cards in a list")
    cards.add_argument("list_id")
    cards.set_defaults(func=list_cards)

    create = subparsers.add_parser("create-card", help="Create a card in a list")
    create.add_argument("list_id")
    create.add_argument("name")
    create.add_argument("--description", "-d", default="")
    create.add_argument("--due", help="Due date accepted by Trello, for example 2026-06-10T12:00:00Z")
    create.set_defaults(func=create_card)

    return parser


def main(argv: list[str] | None = None) -> int:
    load_dotenv()
    parser = build_parser()
    args = parser.parse_args(argv)
    args.func(args)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
