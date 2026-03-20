import os
import datetime
import win32com.client as win32
import sys
import re
from pathlib import Path

import config
import private

# user input for date range
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
                # single date provided?
                single_date = datetime.datetime.strptime(user_input, "%Y%m%d").date()
                return single_date, single_date
                
        except ValueError:
            print(f"{config.colors.BOLD}{config.colors.RED}[ERROR]{config.colors.RESET} Invalid format. Please use YYYYMMDD-YYYYMMDD")

def extract_date_from_filename(filename):
    # look for 8 digit sequence starting with 20 bcs the names all start with the date
    match = re.search(r'(20\d{6})', filename)
    if match:
        try:
            date_str = match.group(1)
            return datetime.datetime.strptime(date_str, "%Y%m%d").date()
        except ValueError:
            pass
    return None

# creates the folders and returns the path object for the specific day folder
def get_output_path(date_obj):
    script_dir = Path(__file__).parent.absolute()
    main_folder = script_dir / "excel_files"
    
    year_str = date_obj.strftime("%Y")  # year
    month_str = date_obj.strftime("%m") # month
    day_str = date_obj.strftime("%d")   # day
    
    target_folder = main_folder / year_str / month_str / day_str
    
    if not target_folder.exists():
        target_folder.mkdir(parents=True, exist_ok=True)
        print(f"{config.colors.BOLD}{config.colors.YELLOW}[STATUS]{config.colors.RESET} Created directory: {target_folder}")
        
    return target_folder

def process_emails(start_date=None, end_date=None, auto_yes=False):
    try:
        if start_date is None or end_date is None:
            start_date, end_date = get_date_range_from_user()
            
        print(f"{config.colors.BOLD}{config.colors.YELLOW}[STATUS]{config.colors.RESET} Searching for emails between {start_date} and {end_date}...")

        # outlook connection
        outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
        
        # list all available accounts (stores)
        stores = outlook.Folders
        selected_store = None
        
        if stores.Count > 1 and not auto_yes:
            print(f"{config.colors.BOLD}{config.colors.CYAN}Available Outlook Accounts:{config.colors.RESET}")
            for i in range(1, stores.Count + 1):
                print(f"  {i}. {stores.Item(i).Name}")
            
            while True:
                try:
                    selection = input(f"{config.colors.BOLD}{config.colors.CYAN}Select account (1-{stores.Count}): {config.colors.RESET}").strip()
                    selection = int(selection)
                    if 1 <= selection <= stores.Count:
                        selected_store = stores.Item(selection)
                        break
                    print(f"{config.colors.BOLD}{config.colors.RED}[ERROR]{config.colors.RESET} Invalid selection.")
                except ValueError:
                    print(f"{config.colors.BOLD}{config.colors.RED}[ERROR]{config.colors.RESET} Please enter a number.")
        else:
            # REMOVE THIS ONCE AUK Y EMAIL CAN BE REMOVED
            if stores.Count >= 2:
                selected_store = stores.Item(2)
            else:
                selected_store = stores.Item(1)
        
        print(f"{config.colors.BOLD}{config.colors.YELLOW}[STATUS]{config.colors.RESET} Using account: {selected_store.Name}")
        
        # find the Inbox folder in the selected account
        inbox = None
        for i in range(1, selected_store.Folders.Count + 1):
            folder = selected_store.Folders.Item(i)
            # inbox names in eng and chinese
            if folder.Name.lower() in ["inbox", "收件匣", "收件箱"]:
                inbox = folder
                break
        
        if inbox is None:
            # try the first folder or raise error
            print(f"{config.colors.BOLD}{config.colors.RED}[ERROR]{config.colors.RESET} Could not find Inbox folder. Available folders:")
            for i in range(1, selected_store.Folders.Count + 1):
                print(f"  - {selected_store.Folders.Item(i).Name}")
            return {}
        
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True) # newest first

        download_count = 0
        downloaded_by_date = {} # return value
        
        for message in messages:
            try:
                subject = message.Subject
                sender = ""
                try: 
                    # sender_email_address can sometimes fail depending on exchange type
                    sender = message.SenderEmailAddress
                except:
                    try:
                        sender = message.Sender.Address
                    except:
                        pass
                
                received_time = message.ReceivedTime
                msg_date = received_time.date()

                print(f"{config.colors.BOLD}{config.colors.BLUE}[SCANNING]{config.colors.RESET} \n {config.colors.MAGENTA}Date={config.colors.RESET}{msg_date} \n {config.colors.MAGENTA}Sender={config.colors.RESET}{sender} \n {config.colors.MAGENTA}Subject={config.colors.RESET}{subject}") 

                # stop if we go past the start date
                if msg_date < start_date:
                    break
                
                # check if within range
                if not (start_date <= msg_date <= end_date):
                    continue
                
                target_directory = get_output_path(msg_date)
                
                matched_source = None
                for source in private.SOURCES:
                    req_subject = source.get("subject_keyword", "").lower()
                    req_sender = source.get("sender_address", "").lower()
                    req_exclude = source.get("exclude_keyword", "")
                    
                    # if a field is blank in the private.py file, ignore
                    # both must match if both fields are present
                    
                    # check exclusion
                    if req_exclude:
                        if isinstance(req_exclude, list):
                            if any(ex.lower() in subject.lower() for ex in req_exclude):
                                continue 
                        elif req_exclude.lower() in subject.lower():
                            continue 

                    # check keyword
                    if req_subject:
                        subject_ok = req_subject in subject.lower()
                        
                    sender_ok = True
                    if req_sender:
                        sender_ok = req_sender in sender.lower()
                    
                    # if both checks passed or were skipped then its a match
                    if subject_ok and sender_ok:
                        matched_source = source
                        break
                
                if matched_source:
                    print(f"{config.colors.BOLD}{config.colors.GREEN}[SUCCESS]{config.colors.RESET} Found match: '{subject}' from {matched_source['name']}")
                    
                    attachments = message.Attachments
                    if attachments.Count > 0:
                        for i in range(1, attachments.Count + 1):
                            attachment = attachments.Item(i)
                            filename = attachment.FileName
                            
                            # check specific attachment exclusion
                            req_att_exclude = matched_source.get("attachment_exclude_keyword", "")
                            if req_att_exclude:
                                if isinstance(req_att_exclude, list):
                                    if any(ex.lower() in filename.lower() for ex in req_att_exclude):
                                        print(f"{config.colors.BOLD}{config.colors.CYAN}[INFO]{config.colors.RESET} Skipping excluded attachment: {filename}")
                                        continue
                                elif req_att_exclude.lower() in filename.lower():
                                    print(f"{config.colors.BOLD}{config.colors.CYAN}[INFO]{config.colors.RESET} Skipping excluded attachment: {filename}")
                                    continue

                            # skip embedded images in mail
                            # remove this check if ur expecting relevant images in the email
                            if filename.lower().endswith(('.png', '.jpg', '.gif', '.bmp')):
                                continue
                            
                            # determine the target directory based on filename date
                            file_date = extract_date_from_filename(filename)
                            if file_date:
                                final_target_directory = get_output_path(file_date)
                            else:
                                # if theres none then just fallback to email received date
                                final_target_directory = get_output_path(msg_date)

                            # save file
                            clean_name = f"{matched_source['name']}_{filename}"
                            save_path = final_target_directory / clean_name
                            
                            if save_path.exists():
                                print(f"{config.colors.BOLD}{config.colors.CYAN}[INFO]{config.colors.RESET} Skipping duplicate: {clean_name}")
                                continue

                            # save individual attachments bcs we dont wanna unzip later
                            print(f"{config.colors.BOLD}{config.colors.YELLOW}[STATUS]{config.colors.RESET} Saving: {save_path}")
                            attachment.SaveAsFile(str(save_path))
                            download_count += 1
                            
                            # add to tracking dict
                            date_str = (file_date or msg_date).strftime("%Y%m%d")
                            if date_str not in downloaded_by_date:
                                downloaded_by_date[date_str] = []
                            downloaded_by_date[date_str].append(str(save_path))
                            
                            # track
                            date_key = msg_date.isoformat()
                            if date_key not in downloaded_by_date:
                                downloaded_by_date[date_key] = []
                            downloaded_by_date[date_key].append(str(save_path))

            except Exception as error:
                print(f"{config.colors.BOLD}{config.colors.RED}[ERROR]{config.colors.RESET} Error processing message: {error}")
                continue

        print(f"{config.colors.BOLD}{config.colors.GREEN}[SUCCESS]{config.colors.RESET} Downloaded {download_count} files.")
        return downloaded_by_date

    except Exception as error:
        print(f"{config.colors.BOLD}{config.colors.RED}[ERROR]{config.colors.RESET} Critical Error connecting to Outlook: {error}")
        return {}

if __name__ == "__main__":
    process_emails()
