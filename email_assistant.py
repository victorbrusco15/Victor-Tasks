#!/usr/bin/env python3
"""
Email Assistant: Summarize emails and suggest responses based on your email style.
Supports Outlook via IMAP and uses Claude AI for intelligent analysis.
"""

import os
import imaplib
import email
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime, timedelta
import anthropic
from typing import Optional
import json


class OutlookEmailAssistant:
    def __init__(
        self,
        email_address: str,
        app_password: str,
        imap_server: str = "outlook.office365.com",
    ):
        """
        Initialize the email assistant.

        Args:
            email_address: Your Outlook email address
            app_password: App password (NOT your regular password)
                         Get it from: https://account.microsoft.com/account/manage-my-microsoft-account
            imap_server: IMAP server (default: Outlook)
        """
        self.email_address = email_address
        self.app_password = app_password
        self.imap_server = imap_server
        self.client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))

    def connect(self) -> imaplib.IMAP4_SSL:
        """Connect to Outlook via IMAP."""
        mail = imaplib.IMAP4_SSL(self.imap_server, 993)
        mail.login(self.email_address, self.app_password)
        return mail

    def get_previous_day_range(self) -> tuple[datetime, datetime]:
        """Get yesterday's date range. On Monday, return Friday-Monday range."""
        today = datetime.now()

        if today.weekday() == 0:  # Monday
            end = today
            start = today - timedelta(days=3)  # Friday
        else:
            end = today
            start = today - timedelta(days=1)  # Yesterday

        return start, end

    def fetch_emails(self, start_date: Optional[datetime] = None, end_date: Optional[datetime] = None) -> list[dict]:
        """
        Fetch emails from Outlook within the date range.

        Args:
            start_date: Start date (default: yesterday)
            end_date: End date (default: today)

        Returns:
            List of email dictionaries with sender, subject, body
        """
        if start_date is None or end_date is None:
            start_date, end_date = self.get_previous_day_range()

        mail = self.connect()
        try:
            mail.select("INBOX")

            # Search for emails in the date range
            search_criteria = f'SINCE {start_date.strftime("%d-%b-%Y")} BEFORE {(end_date + timedelta(days=1)).strftime("%d-%b-%Y")}'
            status, message_ids = mail.search(None, search_criteria)

            if status != "OK" or not message_ids[0]:
                return []

            emails = []
            for msg_id in message_ids[0].split():
                status, msg_data = mail.fetch(msg_id, "(RFC822)")
                if status == "OK":
                    msg = email.message_from_bytes(msg_data[0][1])

                    # Extract email components
                    sender = msg.get("From", "Unknown")
                    subject = msg.get("Subject", "(No subject)")

                    # Extract body
                    body = ""
                    if msg.is_multipart():
                        for part in msg.walk():
                            if part.get_content_type() == "text/plain":
                                body = part.get_payload(decode=True).decode("utf-8", errors="ignore")
                                break
                    else:
                        body = msg.get_payload(decode=True).decode("utf-8", errors="ignore")

                    emails.append({
                        "sender": sender,
                        "subject": subject,
                        "body": body.strip()[:500],  # Limit to 500 chars
                        "date": msg.get("Date"),
                    })

            return emails
        finally:
            mail.close()
            mail.logout()

    def extract_sent_emails_for_pattern(self, limit: int = 20) -> list[dict]:
        """
        Extract recent sent emails to understand your response style.

        Args:
            limit: Number of recent sent emails to analyze

        Returns:
            List of sent emails with subject and body
        """
        mail = self.connect()
        try:
            mail.select("[Gmail]/Sent Mail" if "gmail" in self.imap_server else "Sent Items")

            status, message_ids = mail.search(None, "ALL")
            if status != "OK" or not message_ids[0]:
                return []

            # Get the most recent emails
            message_ids = message_ids[0].split()[-limit:]

            sent_emails = []
            for msg_id in message_ids:
                status, msg_data = mail.fetch(msg_id, "(RFC822)")
                if status == "OK":
                    msg = email.message_from_bytes(msg_data[0][1])

                    body = ""
                    if msg.is_multipart():
                        for part in msg.walk():
                            if part.get_content_type() == "text/plain":
                                body = part.get_payload(decode=True).decode("utf-8", errors="ignore")
                                break
                    else:
                        body = msg.get_payload(decode=True).decode("utf-8", errors="ignore")

                    sent_emails.append({
                        "subject": msg.get("Subject", "(No subject)"),
                        "body": body.strip()[:300],
                    })

            return sent_emails
        finally:
            mail.close()
            mail.logout()

    def analyze_emails(self, emails: list[dict]) -> str:
        """
        Use Claude to analyze emails, summarize topics, and suggest responses.

        Args:
            emails: List of email dictionaries

        Returns:
            Summary and suggested responses as a formatted string
        """
        if not emails:
            return "No emails found in the specified date range."

        # Get your communication style
        sent_emails = self.extract_sent_emails_for_pattern(limit=10)

        style_context = ""
        if sent_emails:
            style_examples = "\n".join([
                f"- Subject: {e['subject']}\n  Response: {e['body'][:200]}..."
                for e in sent_emails[:3]
            ])
            style_context = f"""

**Your typical email response style** (based on recent sent emails):
{style_examples}

Use this style as reference when suggesting responses."""

        # Format emails for Claude
        email_text = "\n\n---\n\n".join([
            f"**From:** {e['sender']}\n**Subject:** {e['subject']}\n**Body:**\n{e['body']}"
            for e in emails
        ])

        # Call Claude
        response = self.client.messages.create(
            model="claude-opus-4-7",
            max_tokens=4096,
            system=f"""You are an email assistant. Your task is to:
1. Summarize the main topics from the received emails
2. Suggest professional responses for each email, matching the user's communication style
3. Organize the output in a clear, actionable format

Keep suggestions concise and practical.{style_context}""",
            messages=[
                {
                    "role": "user",
                    "content": f"""Please analyze these emails and:

1. **SUMMARY**: Extract 3-5 main topics/themes
2. **RESPONSE SUGGESTIONS**: For each email, suggest a response (2-3 sentences) based on my typical style

Here are the emails:

{email_text}

Format the output as:
## SUMMARY OF TOPICS
- Topic 1: ...
- Topic 2: ...

## SUGGESTED RESPONSES
### Email 1: [From and Subject]
[Your suggested response]

### Email 2: [From and Subject]
[Your suggested response]

(and so on...)"""
                }
            ]
        )

        return response.content[0].text

    def send_summary_email(self, summary: str) -> None:
        """
        Send the summary and suggestions via email.

        Args:
            summary: The analysis summary text
        """
        date_range = self._format_date_range()

        msg = MIMEMultipart("alternative")
        msg["Subject"] = f"📧 Email Summary - {date_range}"
        msg["From"] = self.email_address
        msg["To"] = self.email_address

        # Create HTML version
        html = f"""
        <html>
            <head>
                <style>
                    body {{ font-family: Arial, sans-serif; line-height: 1.6; color: #333; }}
                    h2 {{ color: #0078d4; margin-top: 20px; border-bottom: 2px solid #0078d4; }}
                    h3 {{ color: #0078d4; margin-top: 15px; }}
                    pre {{ background: #f5f5f5; padding: 10px; border-radius: 5px; overflow-x: auto; }}
                    .email-block {{ background: #f9f9f9; padding: 15px; margin: 10px 0; border-left: 4px solid #0078d4; }}
                </style>
            </head>
            <body>
                <h1>📧 Your Email Summary</h1>
                <p><strong>Period:</strong> {date_range}</p>

                <div class="email-block">
                    {summary.replace(chr(10), '<br>')}
                </div>

                <hr>
                <p style="color: #666; font-size: 12px;">
                    Generated by Email Assistant at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
                </p>
            </body>
        </html>
        """

        msg.attach(MIMEText(summary, "plain"))
        msg.attach(MIMEText(html, "html"))

        # Send via SMTP
        import smtplib

        with smtplib.SMTP_SSL("smtp-mail.outlook.com", 465) as server:
            server.login(self.email_address, self.app_password)
            server.send_message(msg)

        print(f"✅ Summary sent to {self.email_address}")

    def _format_date_range(self) -> str:
        """Format the date range for display."""
        start, end = self.get_previous_day_range()
        if start.date() == end.date() - timedelta(days=1):
            return start.strftime("%B %d, %Y")
        else:
            return f"{start.strftime('%b %d')} - {end.strftime('%b %d, %Y')}"

    def run(self, send_email: bool = True) -> None:
        """
        Run the complete email analysis workflow.

        Args:
            send_email: Whether to send the summary via email (default: True)
        """
        print("📧 Email Assistant starting...")
        print(f"Fetching emails from {self._format_date_range()}...")

        # Fetch emails
        emails = self.fetch_emails()

        if not emails:
            print("❌ No emails found.")
            return

        print(f"✅ Found {len(emails)} email(s)")
        print("🤖 Analyzing with Claude...")

        # Analyze with Claude
        summary = self.analyze_emails(emails)

        print("\n" + "="*60)
        print(summary)
        print("="*60 + "\n")

        # Send via email
        if send_email:
            print("📤 Sending summary email...")
            self.send_summary_email(summary)
        else:
            print("(Skipped email send)")


def main():
    """Main entry point."""
    # Load credentials from environment variables
    email_address = os.environ.get("OUTLOOK_EMAIL")
    app_password = os.environ.get("OUTLOOK_APP_PASSWORD")

    if not email_address or not app_password:
        print("❌ Missing credentials!")
        print("\nSet these environment variables:")
        print("  export OUTLOOK_EMAIL='your-email@outlook.com'")
        print("  export OUTLOOK_APP_PASSWORD='your-app-password'")
        print("\nTo get an App Password:")
        print("  1. Go to https://account.microsoft.com/account/manage-my-microsoft-account")
        print("  2. Security settings → Advanced security")
        print("  3. App passwords → Generate new app password")
        return

    assistant = OutlookEmailAssistant(email_address, app_password)
    assistant.run(send_email=True)


if __name__ == "__main__":
    main()
