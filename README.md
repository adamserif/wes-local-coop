# Wesleyan Local Co-Op Totals Generator

## Prep
1. Login to Qualtrics, navigate to "My Projects", then select the survey to downloaded responses for (e.g. Local Co-Op Survey Spring 2020).
2. On the top bar, select "Data & Analysis"
3. Go to "Export & Import" -> "Export Data"
4. A window titled "Download Data Table" will pop up, click "Use legacy exporter" in that window."
5. You should be in the CSV tab of the window, click "More options" and select "Use choice text." *Make sure this is the only option that is modified"
6. Download the .CSV, it should be of type "LEGACY_CSV (Full)"

## Running the Script
### Requirements
* Python 3+

Clone the repository to your computer, run:
```
pip install -r requirements.txt
```
Then run:
```
python app.py
```
The application should be running in your browser at 127.0.0.1:5000. Upload the .CSV and fill out the required fields to download the generated .CSV file.
