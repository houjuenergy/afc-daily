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

This script scans the configured Outlook account for daily reports sent by bidders and combines them into a single Excel file.

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
2. Run the installer and make sure you check the box that says **"Add Python to PATH"** during the installation.

## 2. Install Required Programs

Install the required dependencies via the command prompt:

`
pip install -r requirements.txt
`

## 3. Set Up Configuration

If this directory already has a private.py file with the correct bidders, no changes are needed. Otherwise, create a private.py file in this directory and add your email sources following this format:

```python
SOURCES = [
    {
        "name": "",                              # Sender Name
        "subject_keyword": "",                   # Email Subject Match
        "exclude_keyword": "",                   # Exclude Subject Keywords
        "attachment_exclude_keyword": "",        # Exclude Attachments Match
        "sender_address": "COMPANY@example.com", # Email Address Match
    }
]
```

Keep the fields blank if you do not want to specify a filter value.

**Example Configuration:**
```python
SOURCES = [
    {
        "name": "",                             
        "subject_keyword": "AFC daily reports", 
        "exclude_keyword": "",                  
        "attachment_exclude_keyword": "weekly", 
        "sender_address": "@stepower.com",      
    }
]
```

<br>

# How to Use

The program can be executed in **manual** or **auto** mode.

## Manual (Choose your dates)

1. Open Command Prompt.
2. Navigate to the folder containing this program:
`
cd C:\path\to\your\folder
`
3. Execute the script:
`
python main.py
`
4. You will be prompted to choose an Outlook account and enter a date range (YYYYMMDD-YYYYMMDD or press Enter for today).

## Auto (Last 14 days)

1. Open Command Prompt and navigate to the directory.
2. Execute the script with the --auto flag:
`
python main.py --auto
`

<br>

# Where Are My Files?

Original and merged reports are saved in the excel_files/ directory based on the date:

`
excel_files/
  2026/
    03/
      01/
        (original reports here)
        MasterReport_20260301.xlsx (merged report)
`

<br>

# Troubleshooting

**"Could not find inbox folder"**
- Make sure your Outlook application is open and logged into the correct account.

<br>

**"Found X files (Need 16)"**
- A minimum of 16 daily reports are currently required to create a master report.
- To modify this requirement, update Line 31 in main.py (sorry I hardocoded this haha).

`python
if len(valid_files) < 16:
`
Change 16 to the new required amount.

<br>

**Files are skipping as "duplicates"**
- The script prevents duplicate downloads. If you intend to redownload the files, delete the existing ones within the excel_files/ directory first.


If you still have issues, please contact HKE-PM for help (Kin)