# Wesleyan Local Co-Op Script

## Prep
1. Login to Qualtrics, navigate to "My Projects", then select the survey to downloaded responses for (e.g. Local Co-Op Survey Spring 2020).
2. On the top bar, select "Data & Analysis"
3. Go to "Export & Import" -> "Export Data"
4. A window titled "Download Data Table" will pop up, click "Use legacy exporter" in that window."
5. You should be in the CSV tab of the window, click "More options" and select "Use choice text." *Make sure this is the only option that is modified"
6. Download the .CSV, it should be of type "LEGACY_CSV (Full)"
7. Open the .CSV file in a spreadsheet editor, delete the first 11 columns and the 1st row, then export the .CSV as a .XLSX file

## Running the Script
### Requirements
* Python 3+
* openpyxl 3.5

Copy the Python script to your local machine, then run it with the following command:
```
python co-op.py <name of input .xlsx> <name of output .xlsx>
```
For example:
```
python co-op.py input_spring_2020.xlsx totals_spring_2020.xlsx
```
