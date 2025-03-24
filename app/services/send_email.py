import asyncio
import logging
from typing import List

import aiohttp

from app.config.settings import Settings
from app.enum.email_recipient_type import EmailRecipientType
from app.exceptions import EmailSendError


class EmailSender:
    """
    A class to send emails asynchronously using aiohttp.

    Attributes
    ----------
    access_token : str
        The access token for authentication.
    api_scope : str
        The API scope for authentication.
    user_email : str
        The user email for authentication.
    settings : Settings
        The application settings.

    Methods
    -------
    __init__(access_token: str, api_scope: str, user_email: str, settings: Settings):
        Initializes the EmailSender instance with the access token, API scope, user email, and settings.

    async send_emails(bodies: List[str], subjects: List[str], recipients: List[str], cc: List[str], cco: List[str]) -> None:
        Sends emails asynchronously using aiohttp.
    """

    def __init__(
        self, access_token: str, api_scope: str, user_email: str, settings: Settings
    ) -> None:
        """
        Initializes the EmailSender instance with the access token, API scope, user email, and settings.
        """
        self.access_token = access_token
        self.api_scope = api_scope
        self.user_email = user_email
        self.settings = settings

    async def send_emails(
        self,
        bodies: List[str],
        subjects: List[str],
        recipients: List[str],
        cc: List[str],
        cco: List[str],
    ) -> None:
        """
        Sends emails asynchronously using aiohttp.

        Args:
            bodies (List[str]): The list of email bodies.
            subjects (List[str]): The list of email subjects.
            recipients (List[str]): The list of email recipients.
            cc (List[str]): The list of email CC recipients.
            cco (List[str]): The list of email CCO recipients.

        Raises:
            EmailSendError: If there is an error sending the emails.
        """
        async with aiohttp.ClientSession() as session:
            tasks = [
                self._send_email(session, body, subject, recipient, cc[i], cco[i])
                for i, (body, subject, recipient) in enumerate(
                    zip(bodies, subjects, recipients)
                )
            ]
            await asyncio.gather(*tasks)

    async def _send_email(
        self,
        session: aiohttp.ClientSession,
        body: str,
        subject: str,
        recipients: str,
        cc: str,
        cco: str,
    ) -> None:
        """
        Sends a single email using aiohttp.

        Args:
            session (aiohttp.ClientSession): The aiohttp client session.
            body (str): The email body.
            subject (str): The email subject.
            recipients (str): The email recipients, separated by ';'.
            cc (str): The email CC recipients, separated by ';'.
            cco (str): The email CCO recipients, separated by ';'.

        Raises:
            EmailSendError: If there is an error sending the email.
        """
        url = f"{self.settings.GRAPH_API_URL}/users/{self.user_email}/sendMail"
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json",
        }
        to_recipients = self._format_recipients(recipients)
        cc_recipients = self._format_recipients(cc)
        cco_recipients = self._format_recipients(cco)

        payload = {
            "message": {
                "subject": subject,
                "body": {
                    "contentType": "HTML",
                    "content": body,
                },
                EmailRecipientType.TO.value: to_recipients,
            },
            "saveToSentItems": self.settings.SAVE_TO_SENT_ITEMS,
        }
        if cc_recipients:
            payload["message"][EmailRecipientType.CC.value] = cc_recipients
        if cco_recipients:
            payload["message"][EmailRecipientType.BCC.value] = cco_recipients

        logging.info(f"Payload: {payload}")
        try:
            async with session.post(url, headers=headers, json=payload) as response:
                response.raise_for_status()
                logging.info(f"Email sent to {recipients}")
        except Exception as e:
            logging.error(f"Error sending email to {recipients}: {e}")
            raise EmailSendError(f"Error sending email to {recipients}: {e}")

    def _format_recipients(self, recipients: str) -> List[dict]:
        """
        Formats the recipients string into a list of dictionaries.

        Args:
            recipients (str): The recipients string, separated by ';'.

        Returns:
            List[dict]: The formatted recipients.
        """
        return [
            {"emailAddress": {"address": email.strip()}}
            for email in recipients.split(";")
            if email.strip() and email.strip().lower() != "nan"
        ]
