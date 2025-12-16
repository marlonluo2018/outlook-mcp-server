"""Pydantic models for request validation"""

from typing import Optional, List, Union
from pydantic import BaseModel, field_validator, Field


class EmailSearchParams(BaseModel):
    """Parameters for email search operations"""

    search_term: str = Field(..., min_length=1, description="Search term to match")
    days: int = Field(default=7, ge=1, le=30, description="Number of days to look back")
    folder_name: Optional[str] = Field(default=None, description="Folder name to search")
    match_all: bool = Field(default=True, description="Match all terms (AND) or any term (OR)")

    @field_validator("search_term")
    @classmethod
    def validate_search_term(cls, v):
        if not v or not v.strip():
            raise ValueError("search_term must not be empty or whitespace")
        return v.strip()

    @field_validator("folder_name")
    @classmethod
    def validate_folder_name(cls, v):
        if v is not None and v.lower() in ["null", ""]:
            return None
        return v


class EmailListParams(BaseModel):
    """Parameters for listing emails"""

    days: int = Field(default=7, ge=1, le=30, description="Number of days to look back")
    folder_name: Optional[str] = Field(default=None, description="Folder name to list from")

    @field_validator("folder_name")
    @classmethod
    def validate_folder_name(cls, v):
        if v is not None and v.lower() in ["null", ""]:
            return None
        return v


class EmailReplyParams(BaseModel):
    """Parameters for replying to an email"""

    email_number: int = Field(..., ge=1, description="Email's position in cache")
    reply_text: str = Field(..., min_length=1, description="Reply text content")
    to_recipients: Optional[Union[str, List[str]]] = Field(
        default=None, description="To recipients (None preserves original)"
    )
    cc_recipients: Optional[Union[str, List[str]]] = Field(
        default=None, description="CC recipients (None preserves original)"
    )

    @field_validator("reply_text")
    @classmethod
    def validate_reply_text(cls, v):
        if not v or not v.strip():
            raise ValueError("reply_text must not be empty or whitespace")
        return v

    @field_validator("to_recipients", "cc_recipients")
    @classmethod
    def validate_recipients(cls, v):
        if v is None:
            return None

        # Convert single string to list
        if isinstance(v, str):
            if not v.strip():  # If empty or whitespace, treat as None
                return None
            v = [v]

        # Validate each email in list
        if not isinstance(v, list):
            raise ValueError("Recipients must be a string or list of strings")

        # Filter out empty strings and validate remaining emails
        filtered_emails = []
        for email in v:
            if isinstance(email, str) and email.strip():
                filtered_emails.append(email.strip())
            # Skip empty strings and None values silently - don't raise errors

        # Return None if no valid emails remain, otherwise return filtered list
        return filtered_emails if filtered_emails else None

    @field_validator("cc_recipients")
    @classmethod
    def validate_cc_sender_exclusion(cls, v, info):
        """Note: This validator cannot access the original email sender at validation time.
        The actual filtering of original sender from CC will be handled in email_composition.py
        to ensure the original email sender is not included in CC when replying."""
        return v


class EmailComposeParams(BaseModel):
    """Parameters for composing a new email"""

    recipient_email: str = Field(..., min_length=1, description="Recipient email address")
    subject: str = Field(..., min_length=1, description="Email subject")
    body: str = Field(..., min_length=1, description="Email body content")
    cc_email: Optional[str] = Field(default=None, description="CC email address")

    @field_validator("recipient_email", "cc_email")
    @classmethod
    def validate_email(cls, v):
        if v is None:
            return None

        if not v or not v.strip():
            raise ValueError("Email address must not be empty")

        # Basic email validation
        import re

        pattern = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
        if not re.match(pattern, v.strip()):
            raise ValueError(f"Invalid email address format: {v}")

        return v.strip()

    @field_validator("subject", "body")
    @classmethod
    def validate_not_empty(cls, v):
        if not v or not v.strip():
            raise ValueError("Field must not be empty or whitespace")
        return v


class PaginationParams(BaseModel):
    """Parameters for pagination"""

    page: int = Field(default=1, ge=1, description="Page number (1-based)")
    per_page: int = Field(default=5, ge=1, le=50, description="Items per page")


class EmailNumberParam(BaseModel):
    """Parameter for operations requiring an email number"""

    email_number: int = Field(..., ge=1, description="Email number in cache")
