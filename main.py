#!/usr/bin/env python3
import argparse
import sys
import os
import platform
import subprocess
import tempfile

# --- Robust intra-package imports (works as module and as script) ---
try:
    # When run as: python -m outlook_assist_V1.main
    from . import controller
    from .model import Language, GreetingStyle, SignOffStyle
    from .outlook_utils import create_outlook_email
except ImportError:
    # When run directly: python outlook_assist_V1/main.py
    sys.path.append(os.path.dirname(__file__))
    import controller
    from model import Language, GreetingStyle, SignOffStyle
    from outlook_utils import create_outlook_email


def force_utf8_stdio():
    """Ensure standard input/output streams use UTF-8 encoding."""
    try:
        sys.stdin.reconfigure(encoding="utf-8")
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass


def open_system_editor_and_read(initial_text: str = "") -> str:
    """
    Open a system text editor (e.g., Notepad or vi) with optional initial_text, and return the edited text.
    """
    fd, path = tempfile.mkstemp(suffix=".txt", text=True)
    os.close(fd)
    try:
        if initial_text:
            with open(path, "w", encoding="utf-8") as f:
                f.write(initial_text)
        if platform.system() == "Windows":
            subprocess.run(["notepad.exe", path])
        else:
            editor = os.environ.get("EDITOR", "vi")
            subprocess.run([editor, path])
        with open(path, "r", encoding="utf-8") as f:
            return f.read()
    finally:
        try:
            os.remove(path)
        except Exception:
            pass


def interactive_text_input(prompt_hint: str = None) -> str:
    """
    Read multi-line input from the user (for email body content).
    End input by entering a single line with a dot ('.').
    """
    print()
    if prompt_hint:
        print(prompt_hint)
    print("Please paste/type your text. Finish input by entering a single line with a dot (.)")
    lines = []
    try:
        while True:
            line = input()
            if line.strip() == ".":
                break
            lines.append(line)
    except EOFError:
        pass
    return "\n".join(lines)


def parse_args():
    parser = argparse.ArgumentParser(description="Email Assistant CLI - draft and reply to emails with GPT assistance.")
    parser.add_argument("--reply", action="store_true", help="Reply to an incoming email (interactive prompt).")
    parser.add_argument("--greeting", choices=[g.value for g in GreetingStyle], default=GreetingStyle.AUTO.value,
                        help="Greeting style for replies (default: auto)")
    parser.add_argument("--signoff", choices=[s.value for s in SignOffStyle], default=SignOffStyle.BEST_REGARDS.value,
                        help="Sign-off style for replies (default: best_regards)")
    parser.add_argument("--lang", choices=[l.value for l in Language], default=Language.EN.value,
                        help="Language for greeting/sign-off (default: en)")
    return parser.parse_args()


def main():
    force_utf8_stdio()
    args = parse_args()

    # Convert CLI args to enum types
    greet_style = GreetingStyle(args.greeting)
    sign_style = SignOffStyle(args.signoff)
    lang = Language(args.lang)

    print("\n=== Email Assistant (GPT-Powered Email Composer) ===\n")

    if args.reply:
        # Interactive reply flow
        recip_name = input("Your display name (for signing replies) [e.g. Filip]: ").strip() or "Filip"
        sender_name = input("Sender's name (if known, otherwise leave blank): ").strip()
        sender_email = input("Sender's email address (to reply to): ").strip()
        subject = input("Email subject (leave blank if not provided): ").strip()

        print("\nEnter the incoming email body below.")
        incoming_body = interactive_text_input(
            prompt_hint="Paste/type the incoming email content, then enter '.' on a blank line when finished."
        )

        print("\nYou may now provide additional instructions for the reply (tone, key points, etc.), or press Enter for a default concise reply.")
        tone = input("Tone/Style instructions (optional): ").strip()
        extra = interactive_text_input(prompt_hint="Any extra notes for the reply? (finish with '.')")

        try:
            reply_subject, reply_body = controller.generate_reply_email(
                recipient_display_name=recip_name,
                incoming_sender_name=sender_name,
                incoming_sender_email=sender_email,
                incoming_subject=subject,
                incoming_body=incoming_body,
                tone_instructions=tone,
                extra_instructions=extra,
                greeting_style=greet_style,
                signoff_style=sign_style,
                language=lang
            )
        except Exception as e:
            print(f"Error generating reply: {e}")
            return

        # Show the generated reply
        final_to = sender_email or "(no sender email provided)"
        print("\n--- Generated Reply ---")
        print(f"To: {final_to}")
        print(f"Subject: {reply_subject}")
        print(reply_body)

        print("\n--- Opening this reply in Outlook (if available) ---")
        try:
            include_orig = input("Include the original email below the reply? (Y/n): ").strip().lower()
            email_body = reply_body
            if include_orig in ("", "y", "yes"):
                email_body += "\n\n--- Original message ---\n" + incoming_body
            create_outlook_email(final_to, reply_subject, email_body)
        except Exception as e:
            print(f"Outlook integration failed: {e}")
            print("Please copy the reply manually into your email client.")
    else:
        # Interactive email composition flow
        to_addr = input("Recipient email address (or leave blank to decide later): ").strip()
        subject = input("Email subject (or leave blank to decide later): ").strip()
        topic = interactive_text_input(prompt_hint="What is the email about? Describe the content (finish with '.')")

        if not topic:
            print("No email content provided. Exiting.")
            return

        try:
            email_body = controller.generate_new_email(
                recipient_address=to_addr,
                subject_line=subject,
                topic_description=topic
            )
        except Exception as e:
            print(f"Error generating email: {e}")
            return

        print("\n--- Drafted Email ---")
        print(email_body)

        print("\n--- Opening this email in Outlook (if available) ---")
        try:
            create_outlook_email(to_addr, subject, email_body)
        except Exception as e:
            print(f"Outlook integration failed: {e}")
            print("Please copy the email content manually into your email client.")


if __name__ == "__main__":
    main()
