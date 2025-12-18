"""Folder management tools for Outlook MCP Server."""

from typing import Optional
from ..backend.outlook_session import OutlookSessionManager


def move_folder_tool(source_folder_path: str, target_parent_path: str) -> dict:
    """Move a folder and all its emails to a new location.

    Args:
        source_folder_path: Path to the source folder (e.g., "user@company.com/Inbox/SubFolder1")
        target_parent_path: Path to the target parent folder (e.g., "user@company.com/Inbox/NewParent")

    Returns:
        dict: Response containing confirmation message
        {
            "type": "text",
            "text": "Folder moved successfully from 'source_path' to 'target_path' (X emails moved)"
        }

    Note:
        This tool moves the entire folder structure and all emails contained within it.
        Cannot be used to move default folders like Inbox, Sent Items, etc.
        
        IMPORTANT: Folder paths must include the email address as the root folder.
        Use format: "user@company.com/Inbox/SubFolder" not just "Inbox/SubFolder"
    """
    if not source_folder_path or not isinstance(source_folder_path, str):
        raise ValueError("Source folder path must be a non-empty string")
    if not target_parent_path or not isinstance(target_parent_path, str):
        raise ValueError("Target parent path must be a non-empty string")

    try:
        with OutlookSessionManager() as outlook_session:
            result = outlook_session.move_folder(source_folder_path, target_parent_path)
            return {"type": "text", "text": result}
    except Exception as e:
        return {"type": "text", "text": f"Error moving folder: {str(e)}"}


def get_folder_list_tool() -> dict:
    """Lists all Outlook mail folders in a hierarchical structure.

    Returns:
        dict: MCP response with formatted hierarchical folder structure

    """
    try:
        with OutlookSessionManager() as outlook_session:
            result = []
            # Build hierarchy
            for folder in outlook_session.outlook_namespace.Folders:
                result.append(folder.Name)  # Email account level
                result.extend(_get_subfolder_lines(folder, "  "))
            return {"type": "text", "text": "\n".join(result)}
    except Exception as e:
        return {"type": "text", "text": f"Error listing folders: {str(e)}"}


def _get_subfolder_lines(folder, indent):
    """Recursively get subfolder lines with indentation."""
    lines = []
    try:
        for subfolder in folder.Folders:
            lines.append(f"{indent}{subfolder.Name}")
            lines.extend(_get_subfolder_lines(subfolder, indent + "  "))
    except Exception:
        pass
    return lines


def create_folder_tool(folder_name: str, parent_folder_name: Optional[str] = None) -> dict:
    """Create a new folder in the specified parent folder.

    Args:
        folder_name: Name of the folder to create (e.g., "user@company.com/Inbox/NewFolder" or just "NewFolder")
        parent_folder_name: Name of the parent folder (optional, defaults to Inbox)

    Returns:
        dict: Response containing confirmation message
        {
            "type": "text",
            "text": "Folder created successfully: folder_path"
        }
        
    Note:
        For nested folder creation, use full path format: "user@company.com/Inbox/ParentFolder/NewFolder"
        For simple folder creation, you can use just the folder name and specify parent_folder_name
    """
    if not folder_name or not isinstance(folder_name, str):
        raise ValueError("Folder name must be a non-empty string")
    
    try:
        with OutlookSessionManager() as outlook_session:
            result = outlook_session.create_folder(folder_name, parent_folder_name)
            return {"type": "text", "text": result}
    except Exception as e:
        return {"type": "text", "text": f"Error creating folder: {str(e)}"}


def remove_folder_tool(folder_name: str) -> dict:
    """Remove an existing folder.

    Args:
        folder_name: Name or path of the folder to remove (supports nested paths like "user@company.com/Inbox/SubFolder1/SubFolder2")

    Returns:
        dict: Response containing confirmation message
        {
            "type": "text",
            "text": "Folder removed successfully"
        }
        
    Note:
        IMPORTANT: Folder paths must include the email address as the root folder.
        Use format: "user@company.com/Inbox/SubFolder" not just "Inbox/SubFolder"
    """
    if not folder_name or not isinstance(folder_name, str):
        raise ValueError("Folder name must be a non-empty string")

    try:
        with OutlookSessionManager() as outlook_session:
            result = outlook_session.remove_folder(folder_name)
            return {"type": "text", "text": result}
    except Exception as e:
        return {"type": "text", "text": f"Error removing folder: {str(e)}"}