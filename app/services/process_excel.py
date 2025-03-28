import logging
from typing import Dict, List

import pandas as pd

from app.config.settings import Settings
from app.enum.excel_columns import ExcelColumns
from app.exceptions import ExcelReadError


class ExcelProcessor:
    """
    A class to process Excel files and extract email bodies, subjects, and recipients.
    """

    def __init__(self, file_path: str, settings: Settings) -> None:
        """
        Initializes the ExcelProcessor instance with the path to the Excel file and settings.
        """
        self.file_path = file_path
        self.settings = settings
        self.xls = None

    def process_excel(self) -> Dict[str, List[str]]:
        """
        Processes the Excel file and extracts email bodies, subjects, and recipients.

        Returns:
            Dict[str, List[str]]: A dictionary with email bodies, subjects, and recipients.

        Raises:
            ExcelReadError: If there is an error reading the Excel file.
        """
        try:
            self.xls = pd.ExcelFile(self.file_path)
            email_data = {
                "bodies": [],
                "subjects": [],
                "recipients": [],
                "cc": [],
                "cco": [],
            }

            df = pd.read_excel(self.xls, sheet_name=self.xls.sheet_names[0])
            df = self._filter_invalid_rows(df)
            email_data["bodies"] = self._extract_email_bodies(df)
            email_data["subjects"] = df[ExcelColumns.SUBJECT.value].astype(str).tolist()
            email_data["recipients"] = (
                df[ExcelColumns.RECIPIENTS.value].astype(str).tolist()
            )
            email_data["cc"] = (
                df.get(ExcelColumns.CC.value, pd.Series([])).astype(str).tolist()
            )
            email_data["cco"] = (
                df.get(ExcelColumns.CCO.value, pd.Series([])).astype(str).tolist()
            )

            return email_data
        except Exception as e:
            logging.error(f"Error reading Excel file: {e}")
            raise ExcelReadError(f"Error reading Excel file: {e}")

    def _filter_invalid_rows(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Filters out rows with invalid values in "CORPO E-MAIL" columns.

        Args:
            df (pd.DataFrame): The DataFrame to filter.

        Returns:
            pd.DataFrame: The filtered DataFrame.
        """
        body_columns = [
            col for col in df.columns if col.startswith(ExcelColumns.BODY_PREFIX.value)
        ]
        for col in body_columns:
            df[col] = df[col].astype(str)  # Convert column values to strings
            df = df[~df[col].str.lower().isin(self.settings.INVALID_VALUES)]
        return df

    def _extract_email_bodies(self, df: pd.DataFrame) -> List[List[str]]:
        """
        Extracts email bodies from the DataFrame.

        Args:
            df (pd.DataFrame): The DataFrame to process.

        Returns:
            List[List[str]]: The list of email bodies.
        """
        body_columns = [
            col for col in df.columns if col.startswith(ExcelColumns.BODY_PREFIX.value)
        ]
        return df[body_columns].astype(str).values.tolist()

    def close(self) -> None:
        """
        Closes the Excel file.
        """
        if self.xls:
            self.xls.close()
