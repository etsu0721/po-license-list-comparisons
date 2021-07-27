# Project Online License List Comparison
## Description
This script finds the differences, by way of a symmetric difference / disjunctive union, between the lists of Project Online and Power BI licensed users maintained by the Office 365 Team and the Project Online Admin Team. A summary of the differences is written to a text file after successful execution: *comparison_summary.txt*. An exact match on user emails, alone, is used as the basis for comparison. The differences for each license type are written to a separate Excel file (**_diffs.xlsx*) in the folder created for script execution. 

## Run instructions
To run *compare-license-lists.py*, follow these steps:
1. Create a new folder in 'OneDrive - Commonwealth of Kentucky\Div-of-Governance-and-Strategy\Project-Mgmt-Branch\License-Lists\'. The folder name should be the run day's date and must be in the following format: YYYY_MM_DD.
2. Save the O365 Project Online license lists (i.e., Excel files) for the following license types inside the folder created in **Step 1**: PowerBI Professional, Project Online Essential, Professional, and Premium. These files come from the O365 Team upon request for an "orphaned account report".
3. Rename each of the files from **Step 2** with the following filename, respectively: PBI, P1, P3, and P5.
4. Download the 'PWA License Tracker.xlsx' from Microsoft Teams (Division of Governance and Strategy > Project Management Branch > Files), and move it from the Downloads folder to the path in **Step 1** (but not inside the newly created folder using the run day's date).
5. At the path in **Step 1** (but not inside the newly created folder using the run day's date) within File Explorer, click the path bar such that the entire path is highlighted. Then type 'cmd' (without quotes) and press Enter. This should open the Command Prompt at the same path as the File Explorer.
6. Type the name of the script, *compare-license-lists.py*, followed by a space and the name of the folder created in **Step 1**, and then press *Enter* on your keyboard. Generally, the command should look something like this: compare-license-lists.py YYYY_MM_DD

## Future work
* Develop a record linkage algorithm which uses other available data elements to more accurately classify match/non-match.
