import shutil
import os
import zipfile
from cuid2 import cuid_wrapper
from fastapi import UploadFile
from typing import Callable


cuid_generator: Callable[[], str] = cuid_wrapper()


def delete_folders(folders_to_delete):
    """Deletes output and uploads folders
    Args:
      folders_to_delete: A list of folder paths to delete.
    """
    for folder_path in folders_to_delete:
        if os.path.exists(folder_path):
            try:
                shutil.rmtree(folder_path)
                print(f"Deleted folder: {folder_path}")
            except OSError as e:
                print(f"Error deleting folder {folder_path}: {e}")
        else:
            print(f"Folder not found: {folder_path}")


async def save_files(customer: str, files: list[UploadFile]):
    folder_name = f"{customer}-{cuid_generator()}"
    upload_folder = os.path.join("uploads", folder_name)
    os.makedirs(upload_folder, exist_ok=True)
    for file in files:
        file_location = os.path.join(upload_folder, file.filename)
        with open(file_location, "wb+") as file_object:
            file_object.write(await file.read())

    return upload_folder