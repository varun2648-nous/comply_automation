import os

# Path to the folder you want to scan
folder_path = r"C:\path\to\your\folder"

# Output text file
output_file = "file_list.txt"

# Get list of files
files = os.listdir(folder_path)

# Filter only files (ignore subfolders)
file_names = [f for f in files if os.path.isfile(os.path.join(folder_path, f))]

# Write to text file
with open(output_file, "w") as f:
    for name in file_names:
        f.write(name + "\n")

print(f"File list saved to {output_file}")