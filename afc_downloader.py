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
    main_folder = Path(private.PATH)
    
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
        
        download_count = 0
        downloaded_by_date = {} # return value
        
        # ignored folders list
        ignored_folders = ["deleted items", "封件刪除", "已删除邮件", "drafts", "草稿", "outbox", "寄件匣", "发件箱", "junk email", "垃圾郵件", "垃圾邮件"]
        
        def scan_folder(folder):
            nonlocal download_count, downloaded_by_date
            
            # skip trashed or unwanted
            if folder.Name.lower() in ignored_folders:
                return
                
            try:
                messages = folder.Items
                messages.Sort("[ReceivedTime]", True) # newest first
                
                # if there are valid items, notify user
                if messages.Count > 0:
                    print(f"{config.colors.BOLD}{config.colors.YELLOW}[STATUS]{config.colors.RESET} Scanning folder: {folder.Name}")
                
            except Exception:
                # some folders cant be accessed or sorted
                pass
            else:
                for message in messages:
                    try:
                        subject = getattr(message, 'Subject', None)
                        if not subject:
                            continue
                            
                        sender = ""
                        try: 
                            sender = message.SenderEmailAddress
                        except:
                            try:
                                sender = message.Sender.Address
                            except:
                                pass
                        
                        received_time = getattr(message, 'ReceivedTime', None)
                        if not received_time:
                            continue
                            
                        msg_date = received_time.date()

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
                            
                            if req_exclude:
                                if isinstance(req_exclude, list):
                                    if any(ex.lower() in subject.lower() for ex in req_exclude):
                                        continue 
                                elif req_exclude.lower() in subject.lower():
                                    continue 

                            if req_subject:
                                subject_ok = req_subject in subject.lower()
                                
                            sender_ok = True
                            if req_sender:
                                sender_ok = req_sender in sender.lower()
                            
                            if subject_ok and sender_ok:
                                matched_source = source
                                break
                        
                        if matched_source:
                            print(f"{config.colors.BOLD}{config.colors.BLUE}[SCANNING]{config.colors.RESET} \n {config.colors.MAGENTA}Date={config.colors.RESET}{msg_date} \n {config.colors.MAGENTA}Sender={config.colors.RESET}{sender} \n {config.colors.MAGENTA}Subject={config.colors.RESET}{subject}") 
                            print(f"{config.colors.BOLD}{config.colors.GREEN}[SUCCESS]{config.colors.RESET} Found match: '{subject}' from {matched_source['name']}")
                            
                            attachments = message.Attachments
                            if attachments.Count > 0:
                                for i in range(1, attachments.Count + 1):
                                    attachment = attachments.Item(i)
                                    filename = attachment.FileName
                                    
                                    req_att_exclude = matched_source.get("attachment_exclude_keyword", "")
                                    if req_att_exclude:
                                        if isinstance(req_att_exclude, list):
                                            if any(ex.lower() in filename.lower() for ex in req_att_exclude):
                                                print(f"{config.colors.BOLD}{config.colors.CYAN}[INFO]{config.colors.RESET} Skipping excluded attachment: {filename}")
                                                continue
                                        elif req_att_exclude.lower() in filename.lower():
                                            print(f"{config.colors.BOLD}{config.colors.CYAN}[INFO]{config.colors.RESET} Skipping excluded attachment: {filename}")
                                            continue

                                    if filename.lower().endswith(('.png', '.jpg', '.gif', '.bmp')):
                                        continue
                                    
                                    file_date = extract_date_from_filename(filename)
                                    if file_date:
                                        final_target_directory = get_output_path(file_date)
                                    else:
                                        final_target_directory = get_output_path(msg_date)

                                    clean_name = f"{matched_source['name']}_{filename}"
                                    save_path = final_target_directory / clean_name
                                    
                                    if save_path.exists():
                                        print(f"{config.colors.BOLD}{config.colors.CYAN}[INFO]{config.colors.RESET} Skipping duplicate: {clean_name}")
                                        continue

                                    print(f"{config.colors.BOLD}{config.colors.YELLOW}[STATUS]{config.colors.RESET} Saving: {save_path}")
                                    attachment.SaveAsFile(str(save_path))
                                    download_count += 1
                                    
                                    date_str = (file_date or msg_date).strftime("%Y%m%d")
                                    if date_str not in downloaded_by_date:
                                        downloaded_by_date[date_str] = []
                                    downloaded_by_date[date_str].append(str(save_path))
                                    
                                    date_key = msg_date.isoformat()
                                    if date_key not in downloaded_by_date:
                                        downloaded_by_date[date_key] = []
                                    downloaded_by_date[date_key].append(str(save_path))

                    except Exception as error:
                        # continue looping through messages even if one fails
                        continue

            # recursively scan subfolders
            try:
                for sub_i in range(1, folder.Folders.Count + 1):
                    scan_folder(folder.Folders.Item(sub_i))
            except Exception:
                pass
                
        # scan all root folders in the selected account
        for main_i in range(1, selected_store.Folders.Count + 1):
            scan_folder(selected_store.Folders.Item(main_i))

        print(f"{config.colors.BOLD}{config.colors.GREEN}[SUCCESS]{config.colors.RESET} Downloaded {download_count} files.")
        return downloaded_by_date

    except Exception as error:
        print(f"{config.colors.BOLD}{config.colors.RED}[ERROR]{config.colors.RESET} Critical Error connecting to Outlook: {error}")
        return {}

if __name__ == "__main__":
    process_emails()
