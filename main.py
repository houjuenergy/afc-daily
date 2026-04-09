import os
import sys
import shutil
import datetime
import glob
import traceback
import argparse
from pathlib import Path

# subdirectories to path
script_dir = Path(__file__).parent.absolute()

import config
import private
import afc_downloader
import afc_merger

EXCEL_FILES_ROOT = Path(private.PATH)

# get the date range
def get_date_range_from_user():
    while True:
        user_input = input(f"{config.colors.BOLD}{config.colors.CYAN}Enter date range (YYYYMMDD-YYYYMMDD) or press Enter for today: {config.colors.RESET}").strip()
        
        if not user_input:
            today = datetime.datetime.now().date()
            return today, today
            
        try:
            if '-' in user_input:
                start_str, end_str = user_input.split('-')
                start_date = datetime.datetime.strptime(start_str.strip(), "%Y%m%d").date()
                end_date = datetime.datetime.strptime(end_str.strip(), "%Y%m%d").date()
                
                if start_date > end_date:
                    print(f"{config.colors.BOLD}{config.colors.RED}[ERROR]{config.colors.RESET} Start date cannot be after end date.")
                    continue
                    
                return start_date, end_date
            else:
                # single date provided
                single_date = datetime.datetime.strptime(user_input, "%Y%m%d").date()
                return single_date, single_date
                
        except ValueError:
            print(f"{config.colors.BOLD}{config.colors.RED}[ERROR]{config.colors.RESET} Invalid format. Please use YYYYMMDD-YYYYMMDD")

# downloads attachments from Outlook for the range given
def download_attachments(start_date, end_date, auto_yes=False):
    return afc_downloader.process_emails(start_date, end_date, auto_yes=auto_yes)

def process_date_range(start_date, end_date):
    current_date = start_date
    master_reports = []
    
    while current_date <= end_date:
        print(f"\n{config.colors.BOLD}{config.colors.CYAN}Processing: {current_date}...{config.colors.RESET}")
        
        year_str = current_date.strftime("%Y")
        month_str = current_date.strftime("%m")
        day_str = current_date.strftime("%d")
        
        target_folder = EXCEL_FILES_ROOT / year_str / month_str / day_str
        
        if target_folder.exists() and any(target_folder.rglob("*.xlsx")):
            # Check for 16 files requirement
            all_xlsx = list(target_folder.rglob("*.xlsx"))
            valid_files = [f for f in all_xlsx if "MasterReport" not in f.name and not f.name.startswith("~$")]
            
            if len(valid_files) < 16:
                print(f"{config.colors.BOLD}{config.colors.RED}[ERROR]{config.colors.RESET} Skipping merge for {current_date}: Found {len(valid_files)} files (Need 16).")
                current_date += datetime.timedelta(days=1)
                continue

            date_str = current_date.strftime("%Y%m%d")
            output_name = f"MasterReport_{date_str}.xlsx"
            
            print(f"{config.colors.BOLD}{config.colors.YELLOW}[STATUS]{config.colors.RESET} Merging {len(valid_files)} files in {target_folder}...")
            
            try:
                afc_merger.merge_excel_sheets(str(target_folder), output_name)
                
                # check if created
                master_path = target_folder / output_name
                if master_path.exists():
                    # import original sheets
                    afc_merger.import_original_sheets(str(master_path), str(target_folder))
                    
                    print(f"{config.colors.BOLD}{config.colors.GREEN}[SUCCESS]{config.colors.RESET} Master report saved: {master_path}")
                    master_reports.append(str(master_path))
                else:
                     print(f"{config.colors.BOLD}{config.colors.RED}[ERROR]{config.colors.RESET} Failed to create master report.")
            except Exception as e:
                print(f"{config.colors.BOLD}{config.colors.RED}[ERROR]{config.colors.RESET} Merger failed: {e}")
                traceback.print_exc()

        else:
            print(f"{config.colors.BOLD}{config.colors.RED}[ERROR]{config.colors.RESET} Skipping - no files found for {current_date} in {target_folder}")
        
        current_date += datetime.timedelta(days=1)
    
    return master_reports

def main():
    print(f"\n{config.colors.BOLD}{config.colors.CYAN}{'='*50}{config.colors.RESET}")
    print(f"{config.colors.BOLD}{config.colors.CYAN}   AFC BATTERY DAILY REPORT TOOL{config.colors.RESET}")
    print(f"{config.colors.BOLD}{config.colors.CYAN}{'='*50}{config.colors.RESET}\n")
    
    parser = argparse.ArgumentParser(description="AFC Battery Report Tool")
    parser.add_argument("--auto", action="store_true", help="Run in auto mode (last 14 days, non-interactive)")
    args = parser.parse_args()

    try:
        if args.auto:
            print(f"{config.colors.BOLD}{config.colors.MAGENTA}[AUTO MODE]{config.colors.RESET} Running automated check for the last 14 days.")
            end_date = datetime.datetime.now().date()
            start_date = end_date - datetime.timedelta(days=14)
            auto_yes = True
        else:
            # grab that user input for the date range
            start_date, end_date = get_date_range_from_user()
            auto_yes = False
        
        # download attachments
        downloaded_files = download_attachments(start_date, end_date, auto_yes=auto_yes)
        
        if not downloaded_files:
            print(f"{config.colors.BOLD}{config.colors.YELLOW}[STATUS]{config.colors.RESET} No new files were downloaded. Checking for existing files...")
        
        # process and create master report
        master_reports = process_date_range(start_date, end_date)
        
        print(f"\n{config.colors.BOLD}{config.colors.CYAN}{'='*50}{config.colors.RESET}")
        print(f"{config.colors.BOLD}{config.colors.GREEN}   COMPLETED!{config.colors.RESET}")
        print(f"{config.colors.BOLD}{config.colors.CYAN}{'='*50}{config.colors.RESET}")
        
        if master_reports:
            print(f"\n{config.colors.BOLD}{config.colors.YELLOW}Master Reports Created:{config.colors.RESET}")
            for report in master_reports:
                print(f"  - {report}")
        else:
            print(f"\n{config.colors.BOLD}{config.colors.YELLOW}No master reports were created.{config.colors.RESET}")
        
    except Exception as error:
        print(f"\n{config.colors.BOLD}{config.colors.RED}[ERROR]{config.colors.RESET} {error}")
        traceback.print_exc()
    
    print(f"{config.colors.GREEN}Done.{config.colors.RESET}")
    if not args.auto:
        input(f"{config.colors.CYAN}Press Enter to exit.{config.colors.RESET}")

if __name__ == "__main__":
    main()
