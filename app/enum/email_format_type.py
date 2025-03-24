from enum import Enum


class EmailFormatType(Enum):
    """
    Enum representing the different types of email formatting.
    """

    NEGRITO = "Negrito"
    ITALICO = "Itálico"
    SUBLINHADO = "Sublinhado"
    AUMENTAR_FONTE = "Aumentar Fonte"
    HYPERLINK = "hyperlink"
