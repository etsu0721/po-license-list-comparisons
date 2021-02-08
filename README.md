# Run instructions
To run *compare-license-lists.py*, follow these steps:
1. Create a new folder in 'OneDrive - Commonwealth of Kentucky\Division of Governance and Strategy\Project Management Branch\License Lists\'. The folder should be today's date and must be in the following format: YYYY_MM_DD.
2. Save the O365 Project Online license lists (i.e., Excel files) for the following license types inside the folder created in Step 1: PowerBI Professional, Project Online Essential, Professional, and Premium. These files come from the O365 team upon request for an "orphaned account report".
3. Rename each of the files from Step 2 with the following file name, respectively: PBI, P1, P3, and P5.
4. Download the 'PWA License Tracker.xlsx' from Microsoft Teams (Division of Governance and Strategy > Project Management Branch > Files), and move it from the Downloads folder to the path in Step 1 (but not inside the newly created folder using today's date).
5. At the path in Step 1 (but not inside the newly created folder using today's date) within File Explorer, click the path bar such that the entire path is highlighted. Then type 'cmd' (without quotes) and press Enter. This should open the Command Prompt at the same path as the File Explorer.
6. Type the name of the script, *compare-license-lists.py*, followed by a space and the name of the folder created in **Step 1**, and then press **Enter**. Generally, the command should look something like this: compare-license-lists.py YYYY_MM_DD

# Description
This script finds the differences between the lists of Project Online and Power BI licensed users maintained by the Office 365 Team and the Project Management Branch. A summary of the differences are displayed in the terminal after successful running the script. 
A character-by-character match on user emails, alone, is used as the basis for finding differences.
The differences (in both directions) found for each license type are written as separate CSV files to the folder created for the running of the script. Given that there are four (4) lincense types, there are 4 x 2 = 8 files written.

# Future work
* Find intersection of lists and write to a separate CSV file for each license type
* Develop a record linkage algorithm which uses other available data elements to classify difference/intersection (i.e., non-match/match).