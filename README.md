<a href="https://www.buymeacoffee.com/razieleinhorn" target="_blank"><img src="https://cdn.buymeacoffee.com/buttons/default-orange.png" alt="Buy Me A Coffee" height="41" width="174"></a>

# Description
This is a project that aims to use files produced by Israeli credit card companies, organizes them in a neat table and provide a basic dashboard (see screenshot below).

![Google DataStudio Screenshot](/screenshot.png)

# Installation
 
1. Create a new folder on Google Drive where you will store the exported reports. Call it however you like (mine is called "Credit Card Exported Files").
   Enter the folder and make sure to copy the folder ID from the URL (should look like https://drive.google.com/drive/folders/<FOLDER_ID>)
2. Make your own copy of the [template sheet file](https://docs.google.com/spreadsheets/d/1cFWcpH2fhjfQh6ziOo9KEUYCI86uA7WOyUWJMKHLTSM/edit#gid=733610508)
Note: make sure the duplicated file is NOT inside the folder created on stage 1.
3. Cope the new spreadhseet file ID (should look like https://docs.google.com/spreadsheets/d/<SHEET_FILE_ID>/...)
4. Inside the sheet file, select "Tools->Script Editor" from the menu
5. Look for the following parameters and paste the relevant ID from the steps above:
ID_ANALYSIS_FILE = '<FOLDER_ID>'
ID_REPORTS_FOLDER = '<SHEET_FILE_ID>'
Make sure to select "File->Save" from the menu and close the window. 

That's it! the file should be ready now.

# Usage

1. Place an exported report inside the folder created on stage 1 of the installation.
Make sure to convert the file to a google sheet format.
1. Inside the sheet file, select "Credit Card Analysis->Check For New Files"
1. Wait for the "DB" sheet to fill up with the recent data
1. Under "Categories" sheet, all the businesses that are unclassified will appear under the "ללא סיווג" column. 
Copy each business to the relevant category. 
Once you do - the business name should disappear from the "ללא סיווג" column and the chosen categoory will appear near every relevant transaction in the "DB " sheet.

# Tips And Tricks
- Instead of manually converting each file to Google Sheet, you may configure your GDrive to automatically convert every office files to Google format. It will save some tedious work 
