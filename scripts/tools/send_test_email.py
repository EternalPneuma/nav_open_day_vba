"""Send a small SMTP test email to check whether a mail path is blocked.

This script intentionally uses only Python's standard library. Provide SMTP
connection details via command-line flags or environment variables; do not put
passwords in this file.
"""

from __future__ import annotations

import argparse
import getpass
import os
import smtplib
import socket
import ssl
import sys
from datetime import datetime
from email.message import EmailMessage


DEFAULT_RECIPIENT = "jygx_gdsyb@bankcomm.com"


def env(name: str, default: str | None = None) -> str | None:
    value = os.environ.get(name)
    if value is None or value == "":
        return default
    return value


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Send a text-only SMTP test email.",
    )
    parser.add_argument("--smtp-host", default=env("SMTP_HOST"), help="SMTP server host, or SMTP_HOST.")
    parser.add_argument(
        "--smtp-port",
        type=int,
        default=int(env("SMTP_PORT", "25") or "25"),
        help="SMTP server port, or SMTP_PORT. Defaults to 25.",
    )
    parser.add_argument("--from-addr", default=env("SMTP_FROM"), help="Sender email address, or SMTP_FROM.")
    parser.add_argument("--to-addr", default=env("SMTP_TO", DEFAULT_RECIPIENT), help="Recipient email address.")
    parser.add_argument("--username", default=env("SMTP_USERNAME"), help="SMTP auth username, or SMTP_USERNAME.")
    parser.add_argument(
        "--password-env",
        default="SMTP_PASSWORD",
        help="Environment variable containing the SMTP password. Defaults to SMTP_PASSWORD.",
    )
    parser.add_argument("--subject", default="REIT全收益数据通道测试", help="Email subject.")
    parser.add_argument(
        "--body",
        default=None,
        help="Email body. Defaults to a short timestamped test message.",
    )
    parser.add_argument("--starttls", action="store_true", help="Use STARTTLS after connecting.")
    parser.add_argument("--ssl", action="store_true", help="Use SMTP over SSL/TLS from connection start.")
    parser.add_argument(
        "--timeout",
        type=int,
        default=30,
        help="Socket timeout in seconds. Defaults to 30.",
    )
    parser.add_argument("--dry-run", action="store_true", help="Print sanitized settings without sending.")
    return parser.parse_args()


def build_message(args: argparse.Namespace) -> EmailMessage:
    if not args.from_addr:
        raise ValueError("缺少发件人地址：请传 --from-addr 或设置 SMTP_FROM。")

    body = args.body
    if body is None:
        body = (
            "这是一封 REIT 全收益数据邮件通道测试文本。\n"
            f"发送时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            f"发送主机：{socket.gethostname()}\n"
            "无需回复。\n"
        )

    msg = EmailMessage()
    msg["From"] = args.from_addr
    msg["To"] = args.to_addr
    msg["Subject"] = args.subject
    msg.set_content(body)
    return msg


def send_message(args: argparse.Namespace, msg: EmailMessage) -> None:
    if not args.smtp_host:
        raise ValueError("缺少 SMTP 服务器：请传 --smtp-host 或设置 SMTP_HOST。")
    if args.ssl and args.starttls:
        raise ValueError("--ssl 和 --starttls 不能同时使用。")

    password = env(args.password_env)
    context = ssl.create_default_context()

    if args.ssl:
        client: smtplib.SMTP = smtplib.SMTP_SSL(
            args.smtp_host,
            args.smtp_port,
            timeout=args.timeout,
            context=context,
        )
    else:
        client = smtplib.SMTP(args.smtp_host, args.smtp_port, timeout=args.timeout)

    with client:
        client.ehlo()
        if args.starttls:
            client.starttls(context=context)
            client.ehlo()

        if args.username:
            if password is None:
                password = getpass.getpass(f"Password for {args.username}: ")
            client.login(args.username, password)

        client.send_message(msg)


def main() -> int:
    args = parse_args()

    try:
        msg = build_message(args)
        if args.dry_run:
            print("DRY RUN: not sending email.")
            print(f"SMTP: {args.smtp_host}:{args.smtp_port}")
            print(f"Security: {'SSL' if args.ssl else 'STARTTLS' if args.starttls else 'plain'}")
            print(f"Auth: {'yes' if args.username else 'no'}")
            print(f"From: {args.from_addr}")
            print(f"To: {args.to_addr}")
            print(f"Subject: {args.subject}")
            return 0

        send_message(args, msg)
    except (OSError, smtplib.SMTPException, ValueError) as exc:
        print(f"发送失败：{exc}", file=sys.stderr)
        return 1

    print(f"发送成功：{args.from_addr} -> {args.to_addr}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
