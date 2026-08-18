import os
import sys
import argparse
from comparator_engine import (
    read_data_file,
    find_target_files,
    compare_datasets,
    export_to_excel,
    SUPPORTED_EXTENSIONS
)


def print_banner():
    print("=" * 70)
    print("        MULTI-FILE DATA COMPARATOR SCRIPT (CSV / EXCEL)")
    print("=" * 70)


def prompt_path(message, must_be_dir=False, must_be_file=False):
    while True:
        path = input(message).strip()
        # Clean quotes if user dragged and dropped file into terminal
        if (path.startswith('"') and path.endswith('"')) or (path.startswith("'") and path.endswith("'")):
            path = path[1:-1]
        path = os.path.abspath(path)

        if not os.path.exists(path):
            print(f"[ERROR] Path does not exist: {path}\nPlease try again.")
            continue

        if must_be_file and not os.path.isfile(path):
            print(f"[ERROR] Expected a file, but got a directory: {path}\nPlease try again.")
            continue

        if must_be_dir and not os.path.isdir(path):
            print(f"[ERROR] Expected a directory, but got a file: {path}\nPlease try again.")
            continue

        return path


def select_key_columns(source_file_path):
    print("\nReading source file headers...")
    source_df = read_data_file(source_file_path)
    columns = list(source_df.columns)

    print("\nSource File Columns:")
    for idx, col in enumerate(columns, start=1):
        print(f"  [{idx}] {col}")

    while True:
        choice = input(
            "\nEnter key column number(s) to match on (e.g. 1 or 1,2 or column names separated by comma): "
        ).strip()
        if not choice:
            print("[ERROR] Selection cannot be empty.")
            continue

        selected_keys = []
        parts = [p.strip() for p in choice.split(',')]
        
        valid = True
        for part in parts:
            if part.isdigit():
                idx = int(part) - 1
                if 0 <= idx < len(columns):
                    selected_keys.append(columns[idx])
                else:
                    print(f"[ERROR] Invalid column number: {part}")
                    valid = False
                    break
            else:
                # Column name specified directly
                matched = [c for c in columns if c.lower() == part.lower()]
                if matched:
                    selected_keys.append(matched[0])
                else:
                    print(f"[ERROR] Column '{part}' not found in source file columns.")
                    valid = False
                    break

        if valid and selected_keys:
            # Deduplicate preserving order
            selected_keys = list(dict.fromkeys(selected_keys))
            print(f"[OK] Selected key column(s) for comparison: {selected_keys}")
            return selected_keys


def run_interactive_mode():
    print_banner()

    # 1. Get Source File
    source_file = prompt_path(
        "\n1. Enter Source File Path (.csv, .xlsx, .xls): ",
        must_be_file=True
    )

    # 2. Get Target Folder or Files
    print("\n2. Target Files Selection:")
    print("   You can enter a folder path containing target files OR a single target file path.")
    target_input = prompt_path("   Enter Target Folder/File Path: ")

    recursive_input = input("   Search subfolders recursively? (y/N): ").strip().lower()
    recursive = recursive_input == 'y'

    target_files = find_target_files(target_input, recursive=recursive)

    # Filter out source file if it is in the target directory
    source_abs = os.path.abspath(source_file)
    target_files = [f for f in target_files if os.path.abspath(f) != source_abs]

    if not target_files:
        print(f"\n[ERROR] No valid target files ({', '.join(SUPPORTED_EXTENSIONS)}) found in target path.")
        sys.exit(1)

    print(f"\n[OK] Found {len(target_files)} target file(s) to compare against:")
    for f in target_files[:10]:
        print(f"   * {os.path.basename(f)}")
    if len(target_files) > 10:
        print(f"   ... and {len(target_files) - 10} more files.")

    # 3. Select Key Columns
    key_columns = select_key_columns(source_file)

    # 4. Output file path
    default_output = os.path.join(os.getcwd(), "comparison_report.xlsx")
    out_input = input(f"\n4. Enter Output Excel File Path [Default: {default_output}]: ").strip()
    if not out_input:
        output_file = default_output
    else:
        output_file = os.path.abspath(out_input)
        if not output_file.endswith('.xlsx'):
            output_file += '.xlsx'

    # 5. Perform Comparison
    print("\n" + "=" * 70)
    print("[WAIT] Comparing data across target files... Please wait.")
    print("=" * 70)

    try:
        results = compare_datasets(source_file, target_files, key_columns)
        saved_path = export_to_excel(results, output_file)

        # Print Terminal Summary
        stats = results['summary_stats']
        print("\n" + "[SUMMARY] COMPARISON RESULTS SUMMARY")
        print("-" * 50)
        print(f"Source File Name       : {stats['Source File Name']}")
        print(f"Total Source Records   : {stats['Total Source Records']}")
        print(f"Target Files Checked   : {stats['Total Target Files Checked']}")
        print(f"Key Column(s) Used     : {stats['Key Columns Used']}")
        print(f"FOUND in Target Files  : {stats['Found in Targets (Matched)']} ({stats['Match Percentage']})")
        print(f"NOT FOUND in Targets   : {stats['Missing in Targets (Unmatched)']}")
        print("-" * 50)

        print(f"\n[OK] Detailed Excel report successfully saved to:\n   {saved_path}\n")

    except Exception as e:
        print(f"\n[ERROR] Comparison Error: {str(e)}")
        sys.exit(1)


def main():
    parser = argparse.ArgumentParser(
        description="Compare records from one source CSV/Excel file against multiple target files."
    )
    parser.add_argument("-s", "--source", help="Path to source CSV/Excel file")
    parser.add_argument("-t", "--target", help="Path to target directory or target file")
    parser.add_argument("-k", "--keys", help="Comma-separated key column names to match on (e.g. 'Phone Number' or 'Account No,Bank')")
    parser.add_argument("-o", "--output", default="comparison_report.xlsx", help="Path to output Excel file (.xlsx)")
    parser.add_argument("-r", "--recursive", action="store_true", help="Search target directory recursively")

    args = parser.parse_args()

    # If any required CLI flag is missing, fallback to interactive mode
    if not args.source or not args.target or not args.keys:
        run_interactive_mode()
        return

    # CLI automated execution
    print_banner()
    source_file = os.path.abspath(args.source)
    target_input = os.path.abspath(args.target)
    key_columns = [k.strip() for k in args.keys.split(',') if k.strip()]
    output_file = os.path.abspath(args.output)
    if not output_file.endswith('.xlsx'):
        output_file += '.xlsx'

    target_files = find_target_files(target_input, recursive=args.recursive)
    source_abs = os.path.abspath(source_file)
    target_files = [f for f in target_files if os.path.abspath(f) != source_abs]

    if not target_files:
        print(f"[ERROR] No valid target files found in: {target_input}")
        sys.exit(1)

    print(f"Source File  : {source_file}")
    print(f"Target Count : {len(target_files)} target file(s)")
    print(f"Key Column(s): {key_columns}")
    print(f"Output File  : {output_file}")
    print("\n[WAIT] Processing...")

    results = compare_datasets(source_file, target_files, key_columns)
    saved_path = export_to_excel(results, output_file)

    stats = results['summary_stats']
    print("\n" + "[SUMMARY] RESULTS")
    print("-" * 50)
    print(f"Total Source Records : {stats['Total Source Records']}")
    print(f"FOUND in Target Files: {stats['Found in Targets (Matched)']} ({stats['Match Percentage']})")
    print(f"NOT FOUND            : {stats['Missing in Targets (Unmatched)']}")
    print("-" * 50)
    print(f"[OK] Excel report saved to: {saved_path}")


if __name__ == "__main__":
    main()
