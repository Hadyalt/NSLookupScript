import pandas as pd
import subprocess
import csv

# Input Excel file and output CSV file
input_file = "gouda.xlsx"
output_file = "nslookup_results.csv"

# Read Excel file
df = pd.read_excel(input_file)

# Dictionary to store results grouped by user_name
user_results = {}

# Cache to avoid duplicate nslookups
nslookup_cache = {}

def run_nslookup(ip):
    if ip in nslookup_cache:
        return nslookup_cache[ip]

    try:
        result = subprocess.run(
            ["nslookup", ip],
            capture_output=True,
            text=True,
            shell=True
        )
        output = result.stdout

        hostname = "Not found"
        for line in output.splitlines():
            if "Name:" in line or "name =" in line.lower():
                hostname = line.split(":")[-1].strip()
                break
    except Exception as e:
        hostname = f"Error: {e}"

    nslookup_cache[ip] = hostname
    return hostname

# Process each row from Excel
for _, row in df.iterrows():
    ip = str(row["source.ip"]).strip()
    user = str(row["user_name"]).strip()

    hostname = run_nslookup(ip)

    if user not in user_results:
        user_results[user] = {"ips": [], "hosts": []}

    user_results[user]["ips"].append(ip)
    user_results[user]["hosts"].append(hostname)

# Save results to CSV
with open(output_file, mode="w", newline="") as csvfile:
    fieldnames = ["user_name", "source.ips", "nslookup_results"]
    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
    writer.writeheader()

    for user, data in user_results.items():
        writer.writerow({
            "user_name": user,
            "source.ips": data["ips"],           # Writes as Python-style list
            "nslookup_results": data["hosts"]    # Writes as Python-style list
        })

print(f"✅ Done! Results saved to {output_file}")