import os
import time
import json
import logging
import subprocess
from datetime import datetime
from win32com.client import Dispatch

# Logger configuration
logging.basicConfig(
    filename="scan_log.txt",
    level=logging.INFO,
    format="%(asctime)s - %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)

def log_event_json(date, time, status, duration):
    # Function to register events in a JSON file
    log_entry = {
        "date": date,
        "time": time,
        "status": status,
        "duration": duration
    }
    with open("scan_log.json", "a") as log_file:
        json.dump(log_entry, log_file)
        log_file.write("\n")

def log_event_txt(message):
    # Function to register events in a text file
    logging.info(message)

def create_virus_folder():
    # Function to create the 'virus' folder if it doesn't exist
    virus_folder = "virus"
    if not os.path.exists(virus_folder):
        os.makedirs(virus_folder)

def move_to_virus_folder(file_path):
    # Function to move infected files to the 'virus' folder
    virus_folder = "virus"
    file_name = os.path.basename(file_path)
    new_path = os.path.join(virus_folder, file_name)
    os.rename(file_path, new_path)
    return new_path

def scan_with_windows_defender():
    try:
        # Execute full virus scan command via command prompt
        start_time = time.time()
        start_message = "Starting full virus scan via command prompt..."
        log_event_json(datetime.now().strftime("%Y-%m-%d"), datetime.now().strftime("%H:%M:%S"), start_message, 0)
        log_event_txt(start_message)
        
        command = 'powershell.exe -ExecutionPolicy Bypass -Command "Start-MpScan -ScanType FullScan"'
        result = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        
        end_time = time.time()
        duration = end_time - start_time

        if result.returncode == 0:
            status = "Full virus scan completed successfully."
        else:
            status = f"An error occurred during the scan: {result.stderr.strip()}"

        log_event_json(datetime.now().strftime("%Y-%m-%d"), datetime.now().strftime("%H:%M:%S"), status, duration)
        log_event_txt(f"{status} Duration: {duration:.2f} seconds.")

        # Move infected files to 'virus' folder if any
        create_virus_folder()
        if result.returncode != 0:
            infected_files = parse_infected_files(result.stdout)
            for file_path in infected_files:
                moved_file = move_to_virus_folder(file_path)
                log_event_json(datetime.now().strftime("%Y-%m-%d"), datetime.now().strftime("%H:%M:%S"),
                               f"File moved to virus folder: {moved_file}", 0)
                log_event_txt(f"File moved to virus folder: {moved_file}")

    except Exception as e:
        error_message = f"Error during the scan: {e}"
        log_event_json(datetime.now().strftime("%Y-%m-%d"), datetime.now().strftime("%H:%M:%S"), error_message, 0)
        log_event_txt(error_message)

def parse_infected_files(scan_output):
    # Function to parse infected files from scan output
    infected_files = []
    lines = scan_output.splitlines()
    for line in lines:
        if "Found infected file" in line:
            file_path = line.split("Found infected file: ")[-1].strip()
            infected_files.append(file_path)
    return infected_files

def main():
    while True:
        scan_with_windows_defender()
        # Wait for 5 minutes before the next scan
        time.sleep(300)

if __name__ == "__main__":
    main()
