import kagglehub
import pandas as pd
import os
import glob

print("Downloading Lending Club data...")
path = kagglehub.dataset_download("wordsforthewise/lending-club")
print(f"Data downloaded to: {path}")

# Find all CSV files in the directory (including subfolders)
csv_files = glob.glob(os.path.join(path, '**', '*.csv'), recursive=True)

# Filter out directories and keep only actual files
csv_files = [f for f in csv_files if os.path.isfile(f)]

if csv_files:
    # Prefer the accepted dataset (more useful)
    preferred = [f for f in csv_files if 'accepted' in f]
    if preferred:
        file_path = preferred[0]
    else:
        file_path = csv_files[0]
    print(f"Reading: {file_path}")
    df = pd.read_csv(file_path, nrows=100000, low_memory=False)
    df.to_csv('real_loan_data.csv', index=False)
    print("Saved 100,000 rows to real_loan_data.csv")
else:
    print("No CSV files found.")