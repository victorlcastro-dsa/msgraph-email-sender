from typing import Dict, List

from app.config.settings import Settings
from app.enum.email_format_type import EmailFormatType


class EmailFormatter:
    """
    A class to format email messages.
    """

    def __init__(self, settings: Settings) -> None:
        """
        Initializes the EmailFormatter instance with settings.
        """
        self.settings = settings

    def format_emails(
        self, email_data: Dict[str, List[str]], formats: Dict[str, Dict[str, str]]
    ) -> Dict[str, List[str]]:
        """
        Formats the email bodies, subjects, and recipients.

        Args:
            email_data (Dict[str, List[str]]): The dictionary containing email bodies, subjects, and recipients.
            formats (Dict[str, Dict[str, str]]): The dictionary containing the formats for each body part.

        Returns:
            Dict[str, List[str]]: The formatted email data.
        """
        formatted_data = {
            "bodies": [
                self._format_body(body_parts, formats)
                for body_parts in email_data["bodies"]
            ],
            "subjects": email_data["subjects"],
            "recipients": email_data["recipients"],
            "cc": email_data["cc"],
            "cco": email_data["cco"],
        }
        return formatted_data

    def _format_body(
        self, body_parts: List[str], formats: Dict[str, Dict[str, str]]
    ) -> str:
        """
        Formats the email body.

        Args:
            body_parts (List[str]): The email body parts.
            formats (Dict[str, Dict[str, str]]): The format information.

        Returns:
            str: The formatted email body.
        """
        formatted_body = ""
        for i, body in enumerate(body_parts):
            format_info = formats.get(f"CORPO E-MAIL {i + 1}", {})
            font_size = self.settings.DEFAULT_FONT_SIZE  # Default font size
            if format_info.get("formats", {}).get(EmailFormatType.AUMENTAR_FONTE.value):
                font_size += self.settings.FONT_SIZE_INCREMENT  # Increment font size
            if format_info.get("formats", {}).get(EmailFormatType.NEGRITO.value):
                body = f"<b>{body}</b>"
            if format_info.get("formats", {}).get(EmailFormatType.ITALICO.value):
                body = f"<i>{body}</i>"
            if format_info.get("formats", {}).get(EmailFormatType.SUBLINHADO.value):
                body = f"<u>{body}</u>"
            if format_info.get(EmailFormatType.HYPERLINK.value):
                body = f"<a href='{body}'>{body}</a>"
            formatted_body += f"<span style='font-size: {font_size}em;'>{body}</span>"
            line_breaks = format_info.get("line_breaks", 0)
            formatted_body += "<br>" * line_breaks
        return formatted_body.strip()
