```text
 █████╗ ███████╗ ██████╗    ██████╗ ███████╗██████╗  ██████╗ ██████╗ ████████╗
██╔══██╗██╔════╝██╔════╝    ██╔══██╗██╔════╝██╔══██╗██╔═══██╗██╔══██╗╚══██╔══╝
███████║█████╗  ██║         ██████╔╝█████╗  ██████╔╝██║   ██║██████╔╝   ██║   
██╔══██║██╔══╝  ██║         ██╔══██╗██╔══╝  ██╔═══╝ ██║   ██║██╔══██╗   ██║   
██║  ██║██║     ╚██████╗    ██║  ██║███████╗██║     ╚██████╔╝██║  ██║   ██║   
╚═╝  ╚═╝╚═╝      ╚═════╝    ╚═╝  ╚═╝╚══════╝╚═╝      ╚═════╝ ╚═╝  ╚═╝   ╚═╝   
                                                                              
███╗   ███╗███████╗██████╗  ██████╗ ███████╗██████╗                           
████╗ ████║██╔════╝██╔══██╗██╔════╝ ██╔════╝██╔══██╗                          
██╔████╔██║█████╗  ██████╔╝██║  ███╗█████╗  ██████╔╝                          
██║╚██╔╝██║██╔══╝  ██╔══██╗██║   ██║██╔══╝  ██╔══██╗                          
██║ ╚═╝ ██║███████╗██║  ██║╚██████╔╝███████╗██║  ██║                          
╚═╝     ╚═╝╚══════╝╚═╝  ╚═╝ ╚═════╝ ╚══════╝╚═╝  ╚═╝                          
```

> **This code is property of Hou Ju Energy Technology Corporation. See [LICENSE](LICENSE) for more details.**

Scans the Outlook app for any daily reports sent by bidders and combines them into one single Excel file.

| Sections |
|----------|
|[Installation](#installation) |
|[How to Use](#how-to-use) |
|[Where Are My Files?](#where-are-my-files) |
|[Troubleshooting](#troubleshooting) |

<br>

# Installation

## 1. Install Python

1. Download Python from [python.org](https://www.python.org/downloads/)
2. Run the installer (**Note:** Make sure to check the box that says "Add Python to PATH" during the installation)
3. Once Python finishes installing, you're good to go.

## 2. Install Required Programs
I've added a `requirements.txt` file in this directory, it includes all the packages that this computer needs to run the program.
1. Open Command Prompt (search "cmd" in Windows)
2. Copy and paste this command, then press enter:
```
pip install -r requirements.txt
```
3. Wait for it to finish installing everything

## 3. Set Up Configuration
**If this directory already has a `private.py` file and we're using the same bidders as of March 2026, then no changes are needed here.** However, if you're downloading this program from Github or changes to the email list are needed, pls do these steps.
1. Create a `private.py` file in this directory
2. Add your email sources following the existing format:
```python
SOURCES = [
    {
        "name": "",                              # NAME OF SENDER HERE
        "subject_keyword": "",                   # EMAIL SUBJECT HERE
        "exclude_keyword": "",                   # EXCLUDE SUBJECT KEYWORDS HERE
        "attachment_exclude_keyword": "",        # EXCLUDE ATTACHMENTS HERE
        "sender_address": "COMPANY@example.com", # EMAIL HERE
    }
]
```
**Keep the fields blank if you don't want to specify a value to check**
>For example, if I want to download all emails from an example company that has the subject "AFC daily reports", I would need to add the following:
```python
SOURCES = [
    {
        "name": "",                             # Because I don't care specifically who emails me
        "subject_keyword": "AFC daily reports", # I want all emails that contain this subject
        "exclude_keyword": "",                  # No need to exclude since we specified a subject already
        "attachment_exclude_keyword": "weekly", # Sometimes they accidentally send us the weekly reports, I want to ignore these
        "sender_address": "@stepower.com",      # I want the daily reports from all STE Power domains, the sender username doesn't matter
    },
    {
        # Next bidder company goes here
        # The format doesn't allow lists
    }
]
```
<br>

# How to Use

You can use this program in **manual** or **auto** mode.

## Manual (Choose your dates)
1. Open Command Prompt (cmd)
2. Navigate to the folder that you downloaded this program to:
```
cd C:\path\to\your\folder
```
3. Run this script:
```
python afcbattery_report_tool.py
```
4. You'll be prompted to choose your Outlook account and enter a date range (`YYYYMMDD-YYYYMMDD` or just press `Enter` for today)

## Auto (Last 14 days)
1. Open Command Prompt (cmd)
2. Navigate to the folder that you downloaded this program to:
```
cd C:\path\to\your\folder
```
3. Run this script with the `--auto` flag:
```
python afcbattery_report_tool.py --auto
```
**Note:** You need to keep the computer open at all times in order for the script to continue running.

<br>

# Where Are My Files?

Original & merged reports are saved in the \excel_files\ folder:
```
C:\path\to\your\folder\excel_files\
```
Files are organized as follows:
```
excel_files/
  2026/
    03/
      01/
        (original reports here)
        MasterReport_20260301.xlsx (merged report)
```
<br>

# Troubleshooting
**"Could not find inbox folder"**
- Make sure your Outlook app is open and logged in
- I made this script while using Outlook 2016. I don't think it affects the newer apps but it could be a point of failure if nothing works.

<br>

**"Found X files (Need 16)"**
- You might see this in the command line after running the script
- I made it so that you need 16 daily reports to create a master report, so this either means that not all bidders sent you the daily report or we lowered the amount of sites that our company manages
- If you want to change the number, you can do it in the code (sorry I hardcoded this haha)
>Line 70 inside `afcbattery_report_tool.py`
```python
if len(valid_files) < 16:
```
Just change the 16 to another number

<br>

**Files are being skipped as "duplicates"**
- The script won't download the same file twice
- Delete old files if you want to re-download them

---

If you still have issues, please contact HKE-PM for help (Kin)