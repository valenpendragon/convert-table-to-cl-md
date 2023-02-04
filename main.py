# Converting Spreadsheets to Campaign Logger format.
import pandas, openpyxl, os

valid_choice = False
while not valid_choice:
    choice = input("Choose campaign and table to convert. Separate by a semicolon: ")

    try:
        campaign, table = choice.split(";")
    except (ValueError, IndexError):
        print("Invalid choice. Please try again:")
        continue

    if campaign == "Fallow Men":
        base_dir = "./RPGs/Campaigns/Fallow Men/"
    elif campaign == "Cidri":
        base_dir = "./RPGs/Campaigns/Cidri (TFT)/"
    else:
        print("Invalid campaign choice. Try again.")
        continue

    if table in ("Plotline", "Loopy Plans"):
        file_dir = "Plots/Campaign-Planning.xlsx"
    elif table in ("Treasure Table"):
        # elif table in ("Treasure Table", "Treasures"):
        file_dir = "Quartermaster/Quartermaster.xlsx"
    else:
        print("Invalid table choice. Try again.")
        continue

    valid_choice = True

filepath = base_dir + file_dir
df = pandas.read_excel(filepath, sheet_name=table, na_filter=False)
output_lines = []
print(df.shape)
for idx in range(df.shape[0]):
    line = ""
    row = df.loc[idx]
    for col in df.columns:
        line = f"{line}|| {row[col]} "
    line = line + '\n'
    output_lines.append(line)

with open("output.txt", "w") as f:
    for line in output_lines:
        print(line)
        f.writelines(line)