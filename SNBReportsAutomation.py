import shutil
from pathlib import Path
import pandas as pd

from datetime import datetime
from zoneinfo import ZoneInfo 

#to open again and adjust font
from openpyxl import load_workbook
from openpyxl.styles import Font, Border, Side

# ==========CONFIG===========
SOURCE_FOLDER = Path(r"C:\Users\wberame\OneDrive - Lexmark International Inc\Desktop\SNB")
DEST_FOLDER = Path(r"C:\Users\wberame\OneDrive - Lexmark International Inc\Desktop\SNB\Clint and Eric")

ContactReport = "AM0165SX_New - Contacts For Account with Row Id.xlsx"
MP8032Report = "MP8032SP - Asset Detail Report - Extended.xlsx"
#==================

def main():
    #create destination folder if it doesn't exist
    DEST_FOLDER.mkdir(parents=True, exist_ok=True)

    #define full file path
    src_1 = SOURCE_FOLDER / ContactReport
    src_2 = SOURCE_FOLDER / MP8032Report

    dest_1 = DEST_FOLDER / ContactReport
    dest_2 = DEST_FOLDER / MP8032Report

    #copy files
    shutil.copy(src_1, dest_1)
    shutil.copy(src_2, dest_2)

    print("Files copied successfully")

if __name__ == "__main__":
    main()

#load excel files
contact_df = pd.read_excel(DEST_FOLDER / ContactReport, header = 1)
asset_df = pd.read_excel(DEST_FOLDER / MP8032Report, header = 1)

#merge (left join)
merge_df = asset_df.merge(
    contact_df[["#Contact.ContactUID", "Contact.WorkPhone"]],
    left_on="Contact User ID",
    right_on="#Contact.ContactUID",
    how = "left"
)

#drop redudant lookup column
merge_df = merge_df.drop(columns = ["#Contact.ContactUID"])

#move workphone column
phone_col = merge_df.pop("Contact.WorkPhone") #remove temporarily
merge_df.insert(23, "Phone#", phone_col) #insert it back after email column

#columns to drop
columns_to_drop = [
    "CHL4", "CHL5", "CHL6", "CHL7", "CHL8",
    "Network Name", "Computer Name", "Network Topology", "Device Tag ServiceTag", "Device Tag Service Partner", "Special Usage", "Stored Date", "Project", "Equipment ID"
]
merge_df = merge_df.drop(columns = columns_to_drop, errors="ignore")

#current date in Eastern Time
eastern = ZoneInfo("America/New_York")
today_str = datetime.now(tz=eastern).strftime("%Y-%m-%d")

#save to excel
output_file = DEST_FOLDER / f"SNB Asset Detail Report {today_str}.xlsx"
merge_df.to_excel(output_file, index = False)
print(f"Merged report saved to: {output_file}")

#load the file padas just saved
wb = load_workbook(output_file)
ws = wb.active

#remove headers borders for row only
no_border = Border(
    left=Side(border_style=None),
    right=Side(border_style=None),
    top=Side(border_style=None),
    bottom=Side(border_style=None)
)
for cell in ws[1]: #header row
    cell.border = no_border

#insert new top row
ws.insert_rows(1)
ws["A1"] = "#LXK UP MPS Asset"
ws["B1"] = "MP8032SP- Asset Detail - Conversion"

#apply arial 10 default excel font
arial10 = Font(name="ARIAL", size = 10)
for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
    for cell in row:
        cell.font = arial10

#save again
wb.save(output_file)
print("Fonts and border applied successfully")