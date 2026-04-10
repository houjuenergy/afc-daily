import os
import sys
import shutil
import datetime
import glob
import traceback
import tkinter as tk
from tkinter import messagebox, scrolledtext, ttk
import threading
import queue
import re
from pathlib import Path

# subdirectories to path
script_dir = Path(__file__).parent.absolute()

import config
import private
import afc_downloader
import afc_merger

EXCEL_FILES_ROOT = Path(private.PATH)

# downloads attachments from Outlook for the range given
def download_attachments(start_date, end_date, auto_yes=False, account_name=None):
    return afc_downloader.process_emails(start_date, end_date, auto_yes=auto_yes, account_name=account_name)

def process_date_range(start_date, end_date):
    current_date = start_date
    master_reports = []
    quantity_req = 16
    
    while current_date <= end_date:
        print(f"\n{config.colors.BOLD}{config.colors.CYAN}Processing: {current_date}...{config.colors.RESET}")
        
        year_str = current_date.strftime("%Y")
        month_str = current_date.strftime("%m")
        day_str = current_date.strftime("%d")
        
        target_folder = EXCEL_FILES_ROOT / year_str / month_str / day_str
        
        if target_folder.exists() and any(target_folder.rglob("*.xlsx")):
            # check for file quantity requirement
            all_xlsx = list(target_folder.rglob("*.xlsx"))
            valid_files = [f for f in all_xlsx if "MasterReport" not in f.name and not f.name.startswith("~$")]
            
            if len(valid_files) < quantity_req:
                print(f"{config.colors.BOLD}{config.colors.RED}[ERROR]{config.colors.RESET} Skipping merge for {current_date}: Found {len(valid_files)} files (Need {quantity_req}).")
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

def gui_main():
    app = tk.Tk()
    
    # icon stuff
    try:
        icon_path = script_dir / "hke.ico"
        if icon_path.exists():
            app.iconbitmap(str(icon_path))
    except Exception:
        pass
        
    # prog title
    app.title("HKE - AFC Report Merger")
    app.geometry("650x450")

    # GUI logging console setup
    log_area = scrolledtext.ScrolledText(app, wrap=tk.WORD, state='disabled', bg="black", fg="lightgray", font=("Consolas", 10))
    
    log_queue = queue.Queue()
    ansi_split = re.compile(r'(\x1B\[[0-9;]*m)')

    # color tags
    log_area.tag_config('code_0', foreground='lightgray')           # RESET
    log_area.tag_config('code_1', font=("Consolas", 10, "bold"))    # BOLD
    log_area.tag_config('code_90', foreground='gray')               # BLACK
    log_area.tag_config('code_91', foreground='#ff5555')          # RED
    log_area.tag_config('code_92', foreground='#55ff55')          # GREEN
    log_area.tag_config('code_93', foreground='#ffff55')          # YELLOW
    log_area.tag_config('code_94', foreground='#5555ff')          # BLUE
    log_area.tag_config('code_95', foreground='#ff55ff')          # MAGENTA
    log_area.tag_config('code_96', foreground='#55ffff')          # CYAN

    class StdoutRedirector:
        def write(self, string):
            log_queue.put(string)
        def flush(self):
            pass

    # redirect console output to the queue
    sys.stdout = StdoutRedirector()
    sys.stderr = StdoutRedirector()

    active_tags = []

    def process_log_queue():
        while not log_queue.empty():
            try:
                msg = log_queue.get_nowait()
                log_area.config(state='normal')
                
                # split by ansi codes
                parts = ansi_split.split(msg)
                for part in parts:
                    if part.startswith('\x1B['):
                        codes = part[2:-1].split(';')
                        for code in codes:
                            if code == '0':
                                active_tags.clear()
                            else:
                                tag = f"code_{code}"
                                if tag not in active_tags:
                                    active_tags.append(tag)
                    elif part:
                        log_area.insert(tk.END, part, tuple(active_tags))
                
                log_area.see(tk.END)
                log_area.config(state='disabled')
            except queue.Empty:
                break
        app.after(100, process_log_queue)

    process_log_queue()

    def run_process_thread(start_date, end_date, auto_yes):
        acc_name = account_var.get()
        # disable the buttons when running script
        account_cb.config(state=tk.DISABLED)
        manual_btn.config(state=tk.DISABLED)
        log_area.config(state='normal')
        log_area.delete(1.0, tk.END)  # clear logs
        log_area.config(state='disabled')

        def task():
            try:
                print(f"\n{config.colors.CYAN}{'='*50}{config.colors.RESET}")
                print(f"{config.colors.BOLD}{config.colors.CYAN}Starting process for {start_date} to {end_date}...{config.colors.RESET}")
                
                # NOTE: using auto_yes=True directly so input() isnt triggered in console
                downloaded_files = download_attachments(start_date, end_date, auto_yes=True, account_name=acc_name)
                if not downloaded_files:
                    print(f"{config.colors.BOLD}{config.colors.YELLOW}[STATUS]{config.colors.RESET} No new files were downloaded. Checking for existing files...")
                
                master_reports = process_date_range(start_date, end_date)
                
                print(f"\n{config.colors.CYAN}{'='*50}{config.colors.RESET}")
                print(f"{config.colors.BOLD}{config.colors.GREEN}   COMPLETED!{config.colors.RESET}")
                
                if master_reports:
                    print(f"{config.colors.BOLD}{config.colors.GREEN}Completed successfully! Created {len(master_reports)} master reports.{config.colors.RESET}")
                    app.after(0, lambda: messagebox.showinfo("Success", f"Created {len(master_reports)} reports!"))
                else:
                    print(f"{config.colors.BOLD}{config.colors.YELLOW}Completed, but no master reports were created.{config.colors.RESET}")
                    app.after(0, lambda: messagebox.showwarning("Notice", "No master reports were created."))
                    
            except Exception as error:
                print(f"\n{config.colors.BOLD}{config.colors.RED}[ERROR]{config.colors.RESET} {error}")
                traceback.print_exc()
                app.after(0, lambda e=error: messagebox.showerror("Error", f"An error occurred:\n{e}"))
            finally:
                # reenable buttons
                app.after(0, lambda: manual_btn.config(state=tk.NORMAL))
                app.after(0, lambda: account_cb.config(state="readonly"))

        threading.Thread(target=task, daemon=True).start()

    placeholder_text = "YYYYMMDD-YYYYMMDD"

    def on_manual():
        user_input = date_entry.get().strip()
        if user_input == placeholder_text or not user_input:
            today = datetime.datetime.now().date()
            start_date, end_date = today, today
        else:
            try:
                if '-' in user_input:
                    start_str, end_str = user_input.split('-')
                    start_date = datetime.datetime.strptime(start_str.strip(), "%Y%m%d").date()
                    end_date = datetime.datetime.strptime(end_str.strip(), "%Y%m%d").date()
                    if start_date > end_date:
                        messagebox.showerror("Error", "Start date cannot be after end date.")
                        return
                else:
                    start_date = datetime.datetime.strptime(user_input, "%Y%m%d").date()
                    end_date = start_date
            except ValueError:
                messagebox.showerror("Error", "Invalid format. Please use YYYYMMDD-YYYYMMDD")
                return
        
        run_process_thread(start_date, end_date, auto_yes=True)

    # ui
    top_frame = tk.Frame(app)
    top_frame.pack(pady=10, fill=tk.X, padx=10)

    tk.Label(top_frame, text="Account:").pack(side=tk.LEFT, padx=(0, 5))
    
    account_var = tk.StringVar()
    accounts = afc_downloader.get_available_accounts()
    account_cb = ttk.Combobox(top_frame, textvariable=account_var, values=accounts, state="readonly")
    if accounts:
        account_cb.current(0)
    account_cb.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 15))

    tk.Label(top_frame, text="Date:").pack(side=tk.LEFT, padx=(0, 5))
    date_entry = tk.Entry(top_frame)
    
    # placeholder logic
    date_entry.insert(0, placeholder_text)
    date_entry.config(fg='grey')
    
    def on_focus_in(event):
        if date_entry.get() == placeholder_text:
            date_entry.delete(0, tk.END)
            date_entry.config(fg='black')

    def on_focus_out(event):
        if not date_entry.get():
            date_entry.insert(0, placeholder_text)
            date_entry.config(fg='grey')

    date_entry.bind("<FocusIn>", on_focus_in)
    date_entry.bind("<FocusOut>", on_focus_out)
    date_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 15))

    manual_btn = tk.Button(top_frame, text="Run Extractor", command=on_manual, width=15, height=1, bg="#e0e0e0")
    manual_btn.pack(side=tk.LEFT)

    # log text area at the bottom
    log_area.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)

    ASCII_LOGO = f"""
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
    """
    print(ASCII_LOGO)
    
    print(f"Select an account, date range, and click Run Extractor.\n")

    app.mainloop()

if __name__ == "__main__":
    gui_main()
