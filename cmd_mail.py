# svui mzoz htya nqnq
import argparse
import getpass
import mimetypes
import os
import smtplib
import sys
from email.message import EmailMessage

SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 465


def build_message(sender, to_list, subject, body, files):
    msg = EmailMessage()
    msg["From"] = sender
    msg["To"] = ", ".join(to_list)
    msg["Subject"] = subject
    msg.set_content(body)

    for path in files:
        if not os.path.isfile(path):
            print("file not found:", path)
            sys.exit(1)

        # guess the file type so the mail client opens it properly
        ctype, encoding = mimetypes.guess_type(path)
        if ctype is None or encoding is not None:
            ctype = "application/octet-stream"
        maintype, subtype = ctype.split("/", 1)

        with open(path, "rb") as f:
            msg.add_attachment(
                f.read(),
                maintype=maintype,
                subtype=subtype,
                filename=os.path.basename(path),
            )
    return msg


def main():
    parser = argparse.ArgumentParser(description="send an email with attachments")
    parser.add_argument("--to", nargs="+", required=True, help="one or more receivers")
    parser.add_argument("--subject", default="(no subject)")
    parser.add_argument("--body", default="")
    parser.add_argument("--files", nargs="*", default=[], help="files to attach")
    parser.add_argument(
        "--pause", action="store_true", help="wait for Enter before closing the console"
    )
    args = parser.parse_args()

    try:
        sender = os.environ.get("MAIL_USER") or input("Your email: ")
        password = os.environ.get("MAIL_PASS") or getpass.getpass("App password: ")

        msg = build_message(sender, args.to, args.subject, args.body, args.files)

        try:
            with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
                server.login(sender, password)
                server.send_message(msg)
            print(
                "mail sent to",
                ", ".join(args.to),
                "with",
                len(args.files),
                "attachment(s)",
            )
        except smtplib.SMTPAuthenticationError:
            print("login failed - check your email and app password")
        except (smtplib.SMTPException, OSError) as e:
            print("could not send mail:", e)
    finally:
        if args.pause:
            input("Press Enter to close this window...")


if __name__ == "__main__":
    main()
