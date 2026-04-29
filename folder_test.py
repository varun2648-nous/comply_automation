import os
from collections import defaultdict
from datetime import datetime

# 🔹 Update this path
FOLDER_PATH = r"C:\Users\varun\Downloads\downloaded files from comply\YellowWood"

total_files = 0
approval_count = 0
doc_count = 0
no_doc_count = 0

# ✅ date -> count (approval only)
approval_date_counts = defaultdict(int)

# ✅ unique doc dates
doc_dates = set()

for file_name in os.listdir(FOLDER_PATH):
    file_path = os.path.join(FOLDER_PATH, file_name)

    if not os.path.isfile(file_path):
        continue

    total_files += 1

    name_lower = file_name.lower()
    name_without_ext = os.path.splitext(file_name)[0]
    parts = [p.strip() for p in name_without_ext.split("-")]

    # Extract date safely
    date = parts[1] if len(parts) >= 2 else None

    is_approval = "approval" in name_lower
    has_doc_word = "doc" in name_lower
    has_no_doc = "no doc" in name_lower

    # ✅ Approval count
    if is_approval:
        approval_count += 1
        if date:
            approval_date_counts[date] += 1

    # ✅ Doc count (exclude Approval ( No Doc ))
    if has_doc_word and not has_no_doc:
        doc_count += 1
        if date:
            doc_dates.add(date)

    # ✅ Approval ( No Doc )
    if is_approval and has_no_doc:
        no_doc_count += 1


# ✅ Approval date analysis
unique_dates = [d for d, c in approval_date_counts.items() if c == 1]
repeated_dates = {d: c for d, c in approval_date_counts.items() if c > 1}

# ✅ OUTPUT
print(f"Total number of files with 'approval'          : {approval_count}")
print(f"Total number of files with 'doc'               : {doc_count}")
print(f"Total number of files in the folder            : {total_files}")
print(f"Count of unique approval files (date-wise)     : {len(unique_dates)}")
print(f"Count of repeated approval files (date-wise)   : {len(repeated_dates)}")
print(f"Count of unique doc files (date-wise)          : {len(doc_dates)}")
print(f"Count of approval files with '( No Doc )'      : {no_doc_count}")


# ✅ Group repeated approval dates by file count
files_count_to_dates = defaultdict(list)

for date, count in repeated_dates.items():
    files_count_to_dates[count].append(date)

print("\nDates with multiple approval files:")
for count in sorted(files_count_to_dates):
    sorted_dates = sorted(
        files_count_to_dates[count],
        key=lambda d: datetime.strptime(d, "%m.%d.%Y")
    )
    print(f"{', '.join(sorted_dates)} → {count} files")