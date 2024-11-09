import os
import time
from pptx import Presentation
from pptx.exc import PackageNotFoundError
import zipfile

def update_author_in_ppt_files(ppt_folder, new_author, include_subfolders):
    for root, dirs, files in os.walk(ppt_folder):
        for filename in files:
            if filename.endswith(".pptx") and not filename.startswith("~$"):
                ppt_path = os.path.join(root, filename)
                
                try:
                    # Get the original modification time
                    original_mod_time = os.path.getmtime(ppt_path)
                    
                    # Open the presentation and change the author
                    presentation = Presentation(ppt_path)
                    core_properties = presentation.core_properties
                    core_properties.author = new_author
                    presentation.save(ppt_path)
                    
                    # Restore the original modification time
                    os.utime(ppt_path, (original_mod_time, original_mod_time))
                except (PackageNotFoundError, KeyError, zipfile.BadZipFile):
                    print(f"Skipping corrupted or invalid file: {ppt_path}")
        
        if not include_subfolders:
            break

# Prompt the user for the directory containing the PowerPoint files
ppt_folder = input("Please enter the path to your PowerPoint files: ")

# Define the new author name
new_author = "Moshiko Nayman"

# Ask if the user wants to include sub-folders
include_subfolders = input("Do you want to include sub-folders? (Yes/No): ").strip().lower() == 'yes'

# Update the author in PowerPoint files
update_author_in_ppt_files(ppt_folder, new_author, include_subfolders)

print("Author updated for all valid PowerPoint files without changing the modification date.")
