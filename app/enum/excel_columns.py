from enum import Enum


class ExcelColumns(Enum):
    """
    Enum representing the columns in the Excel file used for email processing.
    """

    SUBJECT = "E-MAIL ASSUNTO"
    RECIPIENTS = "E-MAIL PARA (separar com ;)"
    CC = "E-MAIL CC (separar com ;)"
    CCO = "E-MAIL CCO (separar com ;)"
    BODY_PREFIX = "CORPO E-MAIL"
