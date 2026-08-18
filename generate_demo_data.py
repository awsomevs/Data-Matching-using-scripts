import os
import pandas as pd

def create_demo_dataset():
    sample_dir = os.path.join(os.getcwd(), "sample_data")
    targets_dir = os.path.join(sample_dir, "targets")
    os.makedirs(targets_dir, exist_ok=True)

    # 1. Create Source Data (e.g. Suspect / Cyber Crime Case Data)
    source_records = [
        {"Complaint ID": "CC-2026-001", "Phone Number": "9876543210", "Account No": "000123456789", "Suspect Name": "Rajesh Kumar", "Risk Score": "High"},
        {"Complaint ID": "CC-2026-002", "Phone Number": "9123456789", "Account No": "000987654321", "Suspect Name": "Amit Sharma", "Risk Score": "Medium"},
        {"Complaint ID": "CC-2026-003", "Phone Number": "9988776655", "Account No": "000554433221", "Suspect Name": "Vikram Singh", "Risk Score": "High"},
        {"Complaint ID": "CC-2026-004", "Phone Number": "9811223344", "Account No": "000443322110", "Suspect Name": "Pooja Verma", "Risk Score": "Critical"},
        {"Complaint ID": "CC-2026-005", "Phone Number": "9700011122", "Account No": "000111222333", "Suspect Name": "Suresh Patel", "Risk Score": "Low"},
        {"Complaint ID": "CC-2026-006", "Phone Number": "9555444333", "Account No": "000333222111", "Suspect Name": "Neha Gupta", "Risk Score": "High"},
    ]
    source_df = pd.DataFrame(source_records)
    source_file = os.path.join(sample_dir, "source_complaints.xlsx")
    source_df.to_excel(source_file, index=False)
    print(f"Created source file: {source_file}")

    # 2. Target File 1: Mumbai Bank Branch CSV
    mumbai_records = [
        {"Account No": "000123456789", "Branch": "Mumbai Main", "Balance": "150000", "City": "Mumbai"},
        {"Account No": "000554433221", "Branch": "Andheri East", "Balance": "850000", "City": "Mumbai"},
        {"Account No": "999888777666", "Branch": "Bandra", "Balance": "20000", "City": "Mumbai"},
    ]
    mumbai_file = os.path.join(targets_dir, "bank_branch_mumbai.csv")
    pd.DataFrame(mumbai_records).to_csv(mumbai_file, index=False)
    print(f"Created target file 1: {mumbai_file}")

    # 3. Target File 2: Delhi Bank Branch Excel (.xlsx)
    delhi_records = [
        {"Account No": "000987654321", "Branch": "Connaught Place", "Balance": "430000", "City": "Delhi"},
        {"Account No": "000443322110", "Branch": "Dwarka", "Balance": "920000", "City": "Delhi"},
    ]
    delhi_file = os.path.join(targets_dir, "bank_branch_delhi.xlsx")
    pd.DataFrame(delhi_records).to_excel(delhi_file, index=False)
    print(f"Created target file 2: {delhi_file}")

    # 4. Target File 3: Telecom CDR Records CSV
    telecom_records = [
        {"Phone Number": "9876543210", "Telecom Provider": "Airtel", "State": "Maharashtra"},
        {"Phone Number": "9700011122", "Telecom Provider": "Jio", "State": "Gujarat"},
    ]
    telecom_file = os.path.join(targets_dir, "telecom_records.csv")
    pd.DataFrame(telecom_records).to_csv(telecom_file, index=False)
    print(f"Created target file 3: {telecom_file}")

    print("\n[OK] Demo dataset generation complete!")

if __name__ == "__main__":
    create_demo_dataset()
