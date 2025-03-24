from enum import Enum


class EmailRecipientType(Enum):
    """
    Enum representing the different types of email recipients.
    """

    TO = "toRecipients"
    CC = "ccRecipients"
    BCC = "bccRecipients"
