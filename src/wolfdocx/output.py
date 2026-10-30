"""Metadata shared by verified DOCX save receipts."""

from dataclasses import dataclass

DOCX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


@dataclass(frozen=True, kw_only=True)
class SaveOutputParameters:
    """The content type and file format written by a document save."""

    content_type: str = DOCX_CONTENT_TYPE
    written_format: str = "docx"
