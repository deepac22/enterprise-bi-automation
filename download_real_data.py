import kagglehub
import pandas as pd
import os

print("Downloading Lending Club data...")

path = kagglehub.dataset_download("wordsforthewise/lending-club")
print(f"Data downloaded to: {path}")

files = os.listdir(path)
csv_files = [f for f in files if f.endswith('.csv')]

if csv_files:
    file_path = os.path.join(path, csv_files[0])
    print(f"Reading: {file_path}")
    df = pd.read_csv(file_path, nrows=100000, low_memory=False)
    df.to_csv('real_loan_data.csv', index=False)
    print("Saved 100,000 rows to real_loan_data.csv")
else:
    print("No CSV files found.")