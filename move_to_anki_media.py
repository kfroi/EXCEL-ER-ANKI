import os
import shutil
import sys
from pathlib import Path

import tkinter as tk
from tkinter import messagebox


def find_anki_collection_media():
    """
    Attempt to find the Anki collection.media directory for the default user on Windows.
    Returns the path if found, else None.
    """
    # Typical Anki2 user profile path
    appdata = os.environ.get('APPDATA')
    if not appdata:
        return None
    anki2_path = Path(appdata) / 'Anki2'
    if not anki2_path.exists():
        return None
    # Look for user profiles
    for profile in anki2_path.iterdir():
        if profile.is_dir():
            media_dir = profile / 'collection.media'
            if media_dir.exists():
                return media_dir
    return None

def move_images_to_anki_media():
    src_dir = Path('media') / 'media'
    if not src_dir.exists():
        print(f"Source directory {src_dir} does not exist.")
        return
    dest_dir = find_anki_collection_media()
    if not dest_dir:
        print("Could not find Anki collection.media directory.")
        return
    # Confirm with user
    root = tk.Tk()
    root.withdraw()
    if not messagebox.askyesno("Move Images", f"Move all images from {src_dir} to {dest_dir}?\nExisting files will be overwritten."):
        print("Operation cancelled.")
        return
    count = 0
    for file in src_dir.iterdir():
        if file.is_file() and file.suffix.lower() in ['.png', '.jpg', '.jpeg', '.gif', '.bmp']:
            shutil.copy2(file, dest_dir / file.name)
            count += 1
    print(f"Moved {count} image(s) to {dest_dir}")

if __name__ == "__main__":
    move_images_to_anki_media()
