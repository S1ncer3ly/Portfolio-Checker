# Student Portfolio Photo Checker

Checks every PPTX / Google Slides file in a Drive folder for photos on
Week 1-5 and "The journey in their own words", and writes a report to a Google Sheet.

## Folder contents
- drive_photo_checker.py  - main script (Drive -> Google Sheet)
- check_photos.py         - simple local checker for .pptx files on your computer
- requirements.txt        - libraries to install
- run.bat / run.sh        - one-click run (IDs already filled in)
- service_account.json    - YOU add this (new key, never share it)

## Steps
1. Create a NEW key (delete the old exposed one):
   https://console.cloud.google.com/iam-admin/serviceaccounts
   -> portfolio-generator@studentportfolio-507613.iam.gserviceaccount.com
   -> Keys -> Add key -> Create new key -> JSON. Rename to service_account.json
   and put it in this folder.
2. Enable APIs:
   https://console.cloud.google.com/apis/library/drive.googleapis.com
   https://console.cloud.google.com/apis/library/sheets.googleapis.com
3. Share the Drive folder (Viewer) and the Google Sheet (Editor) with
   portfolio-generator@studentportfolio-507613.iam.gserviceaccount.com
4. Install libraries:   pip install -r requirements.txt
5. Run: double-click run.bat (Windows) or ./run.sh (Mac/Linux)
6. Open the Google Sheet to see the results.

## Troubleshooting
- 403/404: folder or sheet not shared with the service account, or an API not enabled.
- "Found 0 presentation(s)": files not shared, or in sub-folders you can't see.
- Local test without Google: python drive_photo_checker.py --local file.pptx
