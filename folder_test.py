import os
import re
from collections import defaultdict
from datetime import datetime

# 🔹 Update this path
FOLDER_PATH = r"C:\Users\varun\Downloads\downloaded files from comply\YellowWood"

# =========================
# 🔧 REGEX FOR DATE
# =========================
date_pattern = re.compile(r"\b\d{2}\.\d{2}\.\d{4}\b")

# =========================
# 🔢 COUNTERS
# =========================
total_files = 0
approval_count = 0
doc_count = 0
no_doc_count = 0

approval_date_counts = defaultdict(int)
doc_date_counts = defaultdict(int)
date_to_approval_without_no_doc = defaultdict(int)

# =========================
# 🔍 MAIN LOOP
# =========================
for file_name in os.listdir(FOLDER_PATH):
    file_path = os.path.join(FOLDER_PATH, file_name)

    if not os.path.isfile(file_path):
        continue

    total_files += 1

    name_lower = file_name.lower()

    # ✅ Robust date extraction (works for ALL file types)
    date_match = date_pattern.search(file_name)
    date = date_match.group() if date_match else None

    # Normalize date (prevents mismatch issues)
    if date:
        try:
            date = datetime.strptime(date, "%m.%d.%Y").strftime("%m.%d.%Y")
        except:
            date = None

    is_approval = "approval" in name_lower
    has_doc_word = "doc" in name_lower
    has_no_doc = "no doc" in name_lower

    # =========================
    # ✅ APPROVAL FILES
    # =========================
    if is_approval:
        approval_count += 1

        if date:
            approval_date_counts[date] += 1

            if not has_no_doc:
                date_to_approval_without_no_doc[date] += 1

        if has_no_doc:
            no_doc_count += 1

    # =========================
    # ✅ DOC FILES (ANY FILE TYPE)
    # =========================
    if has_doc_word and not has_no_doc:
        doc_count += 1

        if date:
            doc_date_counts[date] += 1


# =========================
# 📊 DATE ANALYSIS
# =========================
unique_dates = [d for d, c in approval_date_counts.items() if c == 1]
repeated_dates = {d: c for d, c in approval_date_counts.items() if c > 1}

# =========================
# 🧾 OUTPUT SUMMARY
# =========================
print(f"Total number of files with 'approval'          : {approval_count}")
print(f"Total number of files with 'doc'               : {doc_count}")
print(f"Total number of files in the folder            : {total_files}")
print(f"Count of unique approval files (date-wise)     : {len(unique_dates)}")
print(f"Count of repeated approval files (date-wise)   : {len(repeated_dates)}")
print(f"Count of unique doc files (date-wise)          : {len(doc_date_counts)}")
print(f"Count of approval files with '( No Doc )'      : {no_doc_count}")

# =========================
# 📅 GROUP REPEATED DATES
# =========================
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

# =========================
# ✅ INTEGRITY CHECK
# =========================
calculated_approval_total = sum(approval_date_counts.values())

print("\nApproval duplication integrity check:")
print(f"Total approval files counted               : {approval_count}")
print(f"Approval files accounted for (date-wise)   : {calculated_approval_total}")

integrity_pass = approval_count == calculated_approval_total

if integrity_pass:
    print("✅ PASS: All approval files are properly grouped by date.")
else:
    print("❌ FAIL: Some approval files are missing or misdated!")

# =========================
# 📊 FREQUENCY BREAKDOWN
# =========================
print("\nApproval date frequency breakdown:")
for count in sorted(files_count_to_dates):
    num_dates = len(files_count_to_dates[count])
    print(f"{num_dates} date(s) × {count} file(s) = {num_dates * count}")

print(f"\nComputed total from breakdown             : {calculated_approval_total}")

# =========================
# ✅ APPROVAL vs DOC VALIDATION
# =========================
effective_approval = approval_count - no_doc_count

print("\nApproval vs Doc validation:")
print(f"Approval files (excluding 'No Doc') : {effective_approval}")
print(f"Total Doc files                     : {doc_count}")

approval_doc_pass = doc_count >= effective_approval

if approval_doc_pass:
    print("✅ PASS: Every approval (excluding 'No Doc') has at least one Doc file.")
else:
    print("❌ FAIL: Missing Doc files for some approvals!")
    print(f"Shortage: {effective_approval - doc_count}")

# =========================
# ✅ DATE-LEVEL VALIDATION
# =========================
missing_doc_dates = []

for date, approval_needed in date_to_approval_without_no_doc.items():
    doc_available = doc_date_counts.get(date, 0)

    if doc_available < approval_needed:
        missing_doc_dates.append(
            f"{date} (Approval: {approval_needed}, Doc: {doc_available})"
        )

print("\nDate-level validation:")

date_level_pass = len(missing_doc_dates) == 0

if date_level_pass:
    print("✅ PASS: Every approval date has sufficient Doc files.")
else:
    print("❌ FAIL: Missing or insufficient Doc files for these dates:")
    for item in sorted(
        missing_doc_dates,
        key=lambda x: datetime.strptime(x.split()[0], "%m.%d.%Y")
    ):
        print(item)

# =========================
# 🏁 FINAL RESULT
# =========================
print("\n=========================")
print("FINAL VALIDATION RESULT")
print("=========================")

if integrity_pass and approval_doc_pass and date_level_pass:
    print("🎉 PASS: All files downloaded correctly and validated.")
else:
    print("❌ FAIL: Data inconsistency detected. Please review above output.")