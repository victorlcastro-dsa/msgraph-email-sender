import os
from typing import List, Optional

from dotenv import load_dotenv

from app.exceptions import EnvironmentVariableError


class Settings:
    """
    A class used to load and store configuration settings from environment variables.

    Attributes
    ----------
    LOG_LEVEL : str
        The logging level.
    LOG_FORMAT : str
        The logging format.
    CLIENT_ID : str
        The client ID for authentication.
    TENANT_ID : str
        The tenant ID for authentication.
    CLIENT_SECRET : str
        The client secret for authentication.
    USER_EMAIL : str
        The user email for authentication.
    AIOHTTP_LIMIT : int
        The connection limit for aiohttp.
    API_SCOPE : str
        The API scope for authentication.
    EXCEL_FILE_PATH : str
        The path to the Excel file.
    APP_TITLE : str
        The title of the application.
    APP_ICON_PATH : str
        The path to the application icon.
    DEFAULT_FONT_SIZE : float
        The default font size.
    FONT_SIZE_INCREMENT : float
        The increment value for font size.
    INVALID_VALUES : List[str]
        The list of invalid values for filtering rows.
    GRAPH_API_URL : str
        The URL for the Microsoft Graph API.
    SAVE_TO_SENT_ITEMS : str
        The value for saving sent items.

    Methods
    -------
    __init__():
        Initializes the Settings instance and loads environment variables.
    """

    def __init__(self) -> None:
        """
        Initializes the Settings instance and loads environment variables.
        """
        load_dotenv()
        self.LOG_LEVEL: str = self._get_env_var("LOG_LEVEL", "INFO").upper()
        self.LOG_FORMAT: str = self._get_env_var(
            "LOG_FORMAT", "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
        )
        self.CLIENT_ID: str = self._get_env_var("CLIENT_ID")
        self.TENANT_ID: str = self._get_env_var("TENANT_ID")
        self.CLIENT_SECRET: str = self._get_env_var("CLIENT_SECRET")
        self.USER_EMAIL: str = self._get_env_var("USER_EMAIL")
        self.AIOHTTP_LIMIT: int = int(self._get_env_var("AIOHTTP_LIMIT", 10))
        self.API_SCOPE: str = self._get_env_var("API_SCOPE")
        self.EXCEL_FILE_PATH: str = self._get_env_var("EXCEL_FILE_PATH")
        self.APP_TITLE: str = self._get_env_var("APP_TITLE")
        self.APP_ICON_PATH: str = self._get_env_var("APP_ICON_PATH")
        self.DEFAULT_FONT_SIZE: float = float(self._get_env_var("DEFAULT_FONT_SIZE", 1))
        self.FONT_SIZE_INCREMENT: float = float(
            self._get_env_var("FONT_SIZE_INCREMENT", 0.01)
        )
        self.INVALID_VALUES: List[str] = self._get_env_var(
            "INVALID_VALUES", "x,nan,"
        ).split(",")
        self.GRAPH_API_URL: str = self._get_env_var(
            "GRAPH_API_URL", "https://graph.microsoft.com/v1.0"
        )
        self.SAVE_TO_SENT_ITEMS: str = self._get_env_var("SAVE_TO_SENT_ITEMS", "true")

    @staticmethod
    def _get_env_var(name: str, default: Optional[str] = None) -> str:
        """
        Gets an environment variable or returns a default value.

        Args:
            name (str): The name of the environment variable.
            default (Optional[str]): The default value if the environment variable is not set.

        Returns:
            str: The value of the environment variable or the default value.

        Raises:
            EnvironmentVariableError: If the environment variable is not set and no default value is provided.
        """
        value = os.getenv(name, default)
        if value is None:
            raise EnvironmentVariableError(
                f"Environment variable {name} is not set and no default value provided."
            )
        return value
