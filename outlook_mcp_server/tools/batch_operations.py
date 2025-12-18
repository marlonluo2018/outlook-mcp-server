"""Batch operations tools for Outlook MCP Server."""

from ..backend.batch_operations import batch_forward_emails


def batch_forward_email_tool(email_number: int, csv_path: str, custom_text: str = "") -> dict:
    """Forward an email to recipients listed in a CSV file in batches of 500 (Outlook BCC limit).

    This function uses an email from your cache as a template and forwards it to multiple recipients
    from a CSV file. The email is sent via BCC to protect recipient privacy.

    Args:
        email_number: The number of the email in the cache to use as template (1-based)
        csv_path: Path to CSV file containing recipient email addresses in 'email' column
        custom_text: Optional custom text to prepend to the email body

    CSV Format:
        The CSV file must contain a column named 'email' with recipient email addresses.
        Example:
        ```
        email
        user1@example.com
        user2@example.com
        user3@example.com
        ```

    Returns:
        dict: Response containing batch sending results
        {
            "type": "text",
            "text": "Batch sending completed for X recipients in Y batches: [detailed results]"
        }

    Note:
        - Maximum 500 recipients per batch due to Outlook BCC limitations
        - Invalid email addresses in the CSV will be skipped with warnings
        - The email is sent as BCC to protect recipient privacy
        - Recipients will see it as a forwarded email with "FW:" prefix
        - This function forwards existing emails, it does not compose new ones
    """
    if not isinstance(email_number, int) or email_number < 1:
        raise ValueError("Email number must be a positive integer")
    if not csv_path or not isinstance(csv_path, str):
        raise ValueError("CSV path must be a non-empty string")
    if custom_text is not None and not isinstance(custom_text, str):
        raise ValueError("Custom text must be a string or None")

    try:
        result = batch_forward_emails(email_number, csv_path, custom_text)
        return {"type": "text", "text": result}
    except Exception as e:
        return {"type": "text", "text": f"Error in batch forward operation: {str(e)}"}