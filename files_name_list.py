import os

# -------- FOLDER PATHS --------
folder_a = r"path to folder A"
folder_b = r"path to folder B"


def get_file_names(folder_path):
    """Return a set of file names (ignore subfolders)"""
    return {
        f for f in os.listdir(folder_path)
        if os.path.isfile(os.path.join(folder_path, f))
    }


# Get file sets
files_a = get_file_names(folder_a)
files_b = get_file_names(folder_b)

# Counts
count_a = len(files_a)
count_b = len(files_b)

# Differences
only_in_a = files_a - files_b
only_in_b = files_b - files_a

# -------- OUTPUT --------
print("Folder A file count:", count_a)
print("Folder B file count:", count_b)
print("Count difference:", abs(count_a - count_b))

print("\nFiles present in Folder A but NOT in Folder B:")
if only_in_a:
    for file in sorted(only_in_a):
        print(file)
else:
    print("None")

print("\nFiles present in Folder B but NOT in Folder A:")
if only_in_b:
    for file in sorted(only_in_b):
        print(file)
else:
    print("None")
