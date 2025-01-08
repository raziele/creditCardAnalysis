<a href="https://www.buymeacoffee.com/razieleinhorn" target="_blank"><img src="https://cdn.buymeacoffee.com/buttons/default-orange.png" alt="Buy Me A Coffee" height="41" width="174"></a>

# Description
This is a project that aims to use files produced by Israeli credit card companies, organizes them in a neat table and provide a basic dashboard (see screenshot below).

<img src="/imgs/google-studio-dashboard.png" alt="Google DataStudio Screenshot" style="width:70%;">

# Installation
 
1. Create a new folder on Google Drive where you will store the exported reports. Call it however you like (mine is called "Credit Card Exported Files").
   Enter the folder and make sure to copy the FOLDER_ID from the URL (notice where it is in the screenshot below)
   <img src="/imgs/exports-folder.png" alt="Exports folder screenshot" style="width:70%;">

2. Make your own copy of the [template sheet file](https://docs.google.com/spreadsheets/d/1dbRjdAioleE7Nfdc20TusW30sC10efnTOsuDkujoalA/edit?gid=733610508#gid=733610508)
   (Access is manually provided so it may take a few hours to get access)
4. Inside the newly-copied file you should find a new menu next to "Extensions" called "Credit card analysis"
   <img src="/imgs/menu.png" alt="Menu" style="width:70%;">
5. Click "Detect analysis file" and make sure you get a message it was completed successfully
   <img src="/imgs/detect-file.png" alt="Detect analysis file" style="width:70%;">
4. Click "Configure report folder" and paste the FOLDER_ID from step 1 into the input field and click OK.
   <img src="/imgs/folder-input.png" alt="Input folder ID" style="width:70%;">

That's it!

At any point a Google warning might show up saying you are running an unsafe script.
If you want to use this system you should click proceed. 
Concerned? see security clause below.
   <img src="/imgs/google-unsafe.png" alt="Google unsafe warning" style="width:70%;">

# Usage
1. Place an exported report inside the folder created on stage 1 of the installation.
2. Inside the sheet file, select "Credit Card Analysis->Check For New Files"
3. Wait for the "DB" sheet to fill up with the recent data
4. Under "Categories" sheet, all the businesses that are unclassified will appear under the "ללא סיווג" column. 
Copy each business to the relevant category. 
Once you do - the business name should disappear from the "ללא סיווג" column and the chosen categoory will appear near every relevant transaction in the "DB " sheet.

# Security
- The entire codebase is copied with the template file.
  This means that every analysis is processed *in the premise of your google account.*
- No data is sent out, neither to complement processing nor for analytics purposes (the latter might change in the future but nothing for now)
