import logging
from tkinter import Tk, filedialog
from typing import Dict, Optional

from app.auth.authenticator import Authenticator
from app.config.settings import Settings
from app.services.email_formatter import EmailFormatter
from app.services.process_excel import ExcelProcessor
from app.services.send_email import EmailSender


class HomeController:
    """
    The HomeController class handles the main logic for selecting an Excel file,
    processing it, formatting emails, and sending them.

    Attributes:
        selected_file (Optional[str]): The path to the selected Excel file.
        status_message (str): The status message of the current operation.
        settings (Settings): The application settings.
    """

    def __init__(self) -> None:
        """
        Initializes the HomeController instance with default values.
        """
        self.selected_file: Optional[str] = None
        self.status_message: str = ""
        self.settings = Settings()

    def open_file_dialog(self) -> str:
        """
        Opens a file dialog to select an Excel file.

        Returns:
            str: The path to the selected file or a message indicating no file was selected.
        """
        root = Tk()
        root.withdraw()  # Hide the root window
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        root.destroy()
        if file_path:
            self.selected_file = file_path
            return f"Selected file: {file_path}"
        else:
            self.selected_file = None
            return "No file selected"

    async def send_emails(
        self, sender_email: str, formats: Dict[str, Dict[str, str]]
    ) -> None:
        """
        Sends emails based on the data from the selected Excel file.

        Args:
            sender_email (str): The email address of the sender.
            formats (Dict[str, Dict[str, str]]): The dictionary containing the formats for each body part.

        Raises:
            Exception: If there is an error during the email sending process.
        """
        if not self.selected_file:
            self.status_message = "Por favor, selecione um arquivo primeiro"
            return

        if not sender_email:
            self.status_message = "Por favor, insira o email do remetente"
            return

        try:
            access_token = await self._get_access_token()
            email_data = self._process_excel()
            formatted_email_data = self._format_emails(email_data, formats)
            await self._send_emails(access_token, sender_email, formatted_email_data)
            self.status_message = "Emails enviados com sucesso"
        except Exception as e:
            self.status_message = f"Erro: {e}"
        finally:
            logging.info("Fechando o arquivo Excel.")
            self._close_excel()

    async def _get_access_token(self) -> str:
        """
        Acquires an access token using the Authenticator.

        Returns:
            str: The access token.
        """
        authenticator = Authenticator(
            self.settings.CLIENT_ID,
            self.settings.TENANT_ID,
            self.settings.CLIENT_SECRET,
            self.settings.API_SCOPE,
        )
        return await authenticator.get_access_token()

    def _process_excel(self) -> Dict[str, list]:
        """
        Processes the selected Excel file and extracts email data.

        Returns:
            Dict[str, list]: A dictionary containing email bodies, subjects, and recipients.
        """
        excel_processor = ExcelProcessor(self.selected_file, self.settings)
        return excel_processor.process_excel()

    def _format_emails(
        self, email_data: Dict[str, list], formats: Dict[str, Dict[str, str]]
    ) -> Dict[str, list]:
        """
        Formats the email data.

        Args:
            email_data (Dict[str, list]): The raw email data.
            formats (Dict[str, Dict[str, str]]): The dictionary containing the formats for each body part.

        Returns:
            Dict[str, list]: The formatted email data.
        """
        email_formatter = EmailFormatter(self.settings)
        return email_formatter.format_emails(email_data, formats)

    async def _send_emails(
        self, access_token: str, sender_email: str, email_data: Dict[str, list]
    ) -> None:
        """
        Sends the formatted emails using the EmailSender.

        Args:
            access_token (str): The access token for authentication.
            sender_email (str): The email address of the sender.
            email_data (Dict[str, list]): The formatted email data.
        """
        email_sender = EmailSender(
            access_token, self.settings.API_SCOPE, sender_email, self.settings
        )
        await email_sender.send_emails(
            email_data["bodies"],
            email_data["subjects"],
            email_data["recipients"],
            email_data["cc"],
            email_data["cco"],
        )

    def _close_excel(self) -> None:
        """
        Closes the Excel file.
        """
        excel_processor = ExcelProcessor(self.selected_file, self.settings)
        excel_processor.close()
