# What About My Privacy? 

It's pretty simple:
When you use the code in this repository - the code and logic executes entirely inside your Google account.
The code doesn't send any information outside of your personal account and doesn't pull any information from an outside location.

In order for the script to work, it basically needs access to two resources:
1. The folder where you store your credit card reports and the reports themselves.
The script will read all reports, parse and reorganize them in a coherent way.
The required access to the folder and files is read-only.

2. The analysis Google sheet itself, which hosts both the code (via Google App Script) and the parsed information from the reports.
This requires both read and write access to the Google sheet. 


