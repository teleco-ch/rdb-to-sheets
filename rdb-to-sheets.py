import sys
import pandas as pd
import re

# Ensure its easier to use for normies
if len(sys.argv) < 2:
    print("Usage: python rdb-to-sheets.py <input_file> [output_file]")
    print("Example: python rdb-to-sheets.py RDBv2.txt RDB.xlsx")
    print("If no output file is defined RDB.xlsx will be used per default make sure such doesn't exist already")
    sys.exit(1)  # Committ da sus side for script!

# Set input and output file 
input_file = sys.argv[1]
output_path = sys.argv[2] if len(sys.argv) > 2 else "RDB.xlsx"  # Default to RDB.xlsx if user of script doesnt define it lmao


with open(input_file, "r") as file:
    lines = file.readlines()

# Look at the data
def parse_rdb_data(lines):
    entries = []
    current_entry = {}
    current_tns = []

    for line in lines:
        line = line.strip()
        if not line:
            continue

        # New DN entry starts
        if line.startswith("DN"):
            if current_entry:  # How about we save the previous entry before we look at a new one omg
                if current_tns:
                    current_entry["TNs"] = current_tns
                entries.append(current_entry)
                current_tns = []
            current_entry = {"DN": line.split()[1]}

        elif "CPND" in line:
            current_entry["NAME"] = None

        elif "NAME" in line and current_entry.get("NAME") is None:
            current_entry["NAME"] = line.replace("NAME", "").strip()

        elif line.startswith("TYPE"):
            current_entry["TYPE"] = line.split()[1]

        elif line.startswith("ROUT"):
            current_entry["ROUT"] = line.split()[1]

        elif line.startswith("STCD"):
            current_entry["STCD"] = line.split()[1]

        elif line.startswith("FEAT"):
            current_entry["FEAT"] = " ".join(line.split()[1:])

        elif line.startswith("TN"):
            tn_match = re.search(r"(\d{3} \d \d{2} \d{2})", line)
            if tn_match:
                tn_value = tn_match.group(1)
                tn_entry = {"TN_1_TN": tn_value}

                key_match = re.search(r"KEY (\d{2})", line)
                if key_match:
                    tn_entry["KEY"] = key_match.group(1)

                des_match = re.search(r"DES\s+(\S+)", line)
                if des_match:
                    tn_entry["DES"] = des_match.group(1)

                date_match = re.search(r"(\d{1,2} \w{3} \d{4})", line)
                if date_match:
                    tn_entry["DATE"] = date_match.group(0)

                tn_entry["MARP"] = "YES" if "MARP" in line else "NO"
                current_tns.append(tn_entry)

    if current_entry:
        if current_tns:
            current_entry["TNs"] = current_tns
        entries.append(current_entry)

    return entries

parsed_entries = parse_rdb_data(lines)

# Categorize them and shit
def categorize_entries(entries):
    main_entries = []
    multiple_tns_entries = []
    route_entries = []
    cdp_entries = []
    feature_code_entries = []
    att_ldn_entries = []
    other_entries = []
    unexpected_data = []

    for entry in entries:
        entry_type = entry.get("TYPE")

        if entry_type in ["SL1", "500"]:
            if "TNs" in entry and len(entry["TNs"]) > 1:
                for tn in entry["TNs"]:
                    multi_entry = entry.copy()
                    multi_entry.update(tn)
                    multiple_tns_entries.append(multi_entry)
            elif "TNs" in entry:
                for tn in entry["TNs"]:
                    main_entry = entry.copy()
                    main_entry.update(tn)
                    main_entries.append(main_entry)
            else:
                main_entries.append(entry)

        elif entry_type in ["ATT", "LDN"]:
            if "TNs" in entry:
                for tn in entry["TNs"]:
                    att_ldn_entry = entry.copy()
                    att_ldn_entry.update(tn)
                    att_ldn_entries.append(att_ldn_entry)
            else:
                att_ldn_entries.append(entry)

        elif entry_type == "RDB":
            route_entries.append(entry)

        elif entry_type == "CDP":
            cdp_entries.append({"DN": entry.get("DN", ""), "TYPE": entry.get("TYPE", ""), "STCD": entry.get("STCD", "DSC")})

        elif entry_type == "FFC":
            feature_code_entries.append(entry)

        elif entry_type is None:
            unexpected_data.append(entry)

        else:
            other_entries.append(entry)

    unexpected_sheet_name = "No Unexpected Data" if not unexpected_data else "Unexpected Data"

    return main_entries, multiple_tns_entries, route_entries, cdp_entries, feature_code_entries, att_ldn_entries, other_entries, unexpected_data, unexpected_sheet_name

# Categorize entries further
main_entries, multiple_tns_entries, route_entries, cdp_entries, feature_code_entries, att_ldn_entries, other_entries, unexpected_data, unexpected_sheet_name = categorize_entries(parsed_entries)


# Save all the data to excel but dropping the raw unprocessed TNs tab that gets made for some reason I dont know why I hate coding and python
with pd.ExcelWriter(output_path) as writer:
    for sheet_name, data in {
        "Main": main_entries,
        "Multiple TNs": multiple_tns_entries,
        "Routes": route_entries,
        "CDP": cdp_entries,
        "Feature Codes": feature_code_entries,
        "ATT_LDN": att_ldn_entries,
        "Other": other_entries,
        unexpected_sheet_name: unexpected_data
    }.items():
        df = pd.DataFrame(data)
        if "TNs" in df.columns:
            df = df.drop(columns=["TNs"])  # Remove TNs column if present
        df.to_excel(writer, sheet_name=sheet_name, index=False)

print("Processing complete. Output saved to", output_path)
