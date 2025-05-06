"""
File Manager Module for WhatsApp Sender Application

This module provides utility functions for managing files used by the WhatsApp Sender Application.
It handles operations like creating, backing up, and deleting Excel files, text files, and PDF files.
"""

import os
import shutil
import datetime
import pandas as pd

def create_backup_folder():
    """Create a backup folder if it doesn't exist."""
    backup_folder = os.path.join(os.getcwd(), "backups")
    if not os.path.exists(backup_folder):
        os.makedirs(backup_folder)
    return backup_folder

def backup_file(file_path):
    """Create a backup of a file with timestamp in the filename."""
    if not os.path.exists(file_path):
        return None
    
    backup_folder = create_backup_folder()
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    file_name = os.path.basename(file_path)
    base_name, extension = os.path.splitext(file_name)
    backup_file_name = f"{base_name}_{timestamp}{extension}"
    backup_path = os.path.join(backup_folder, backup_file_name)
    
    try:
        shutil.copy2(file_path, backup_path)
        print(f"Backup created: {backup_path}")
        return backup_path
    except Exception as e:
        print(f"Error creating backup: {e}")
        return None

def delete_excel_file(file_path, backup=True):
    """Delete an Excel file, optionally creating a backup first."""
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return False
    
    if backup:
        backup_file(file_path)
    
    try:
        os.remove(file_path)
        print(f"Deleted file: {file_path}")
        return True
    except Exception as e:
        print(f"Error deleting file: {e}")
        return False

def create_empty_excel(file_path, columns=None):
    """Create a new empty Excel file with specified columns."""
    if columns is None:
        columns = ['MOBILE']
    
    try:
        df = pd.DataFrame(columns=columns)
        df.to_excel(file_path, index=False)
        print(f"Created new Excel file: {file_path}")
        return True
    except Exception as e:
        print(f"Error creating Excel file: {e}")
        return False

def delete_text_file(file_path, backup=True):
    """Delete a text file, optionally creating a backup first."""
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return False
    
    if backup:
        backup_file(file_path)
    
    try:
        os.remove(file_path)
        print(f"Deleted file: {file_path}")
        return True
    except Exception as e:
        print(f"Error deleting file: {e}")
        return False

def create_empty_text_file(file_path, content=""):
    """Create a new empty text file with optional default content."""
    try:
        with open(file_path, 'w', encoding='utf-8') as file:
            file.write(content)
        print(f"Created new text file: {file_path}")
        return True
    except Exception as e:
        print(f"Error creating text file: {e}")
        return False

def handle_pdf_file(file_path, action="backup"):
    """Handle PDF files - backup or other actions."""
    if not os.path.exists(file_path):
        print(f"PDF file not found: {file_path}")
        return False
    
    if action == "backup":
        return backup_file(file_path) is not None
    
    return False

def list_files_in_directory(directory_path, file_extension=None):
    """List all files in a directory, optionally filtered by extension."""
    if not os.path.exists(directory_path):
        print(f"Directory not found: {directory_path}")
        return []
    
    files = []
    for file in os.listdir(directory_path):
        file_path = os.path.join(directory_path, file)
        if os.path.isfile(file_path):
            if file_extension is None or file.endswith(file_extension):
                files.append(file_path)
    
    return files

def ensure_directory_exists(directory_path):
    """Ensure that a directory exists, creating it if necessary."""
    if not os.path.exists(directory_path):
        try:
            os.makedirs(directory_path)
            print(f"Created directory: {directory_path}")
            return True
        except Exception as e:
            print(f"Error creating directory: {e}")
            return False
    return True
