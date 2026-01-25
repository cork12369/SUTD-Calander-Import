import pdfplumber
import csv
import re
from datetime import datetime, timedelta
import sys
import os
import argparse
import tkinter as tk
from tkinter import filedialog, messagebox

# Default Output filename
DEFAULT_OUTPUT_CSV = "class_schedule.csv"

# Google Calendar CSV Headers
CSV_HEADERS = [
    "Subject", "Start Date", "Start Time", "End Date", "End Time",
    "All Day Event", "Description", "Location", "Private"
]

def parse_time(time_str):
    """Parses time strings like '10:00AM', '2:30 PM', 'Mon 8:30' into 'HH:MM AM/PM'."""
    if not time_str:
        return None
    
    # Clean string to uppercase for easier matching
    time_str = time_str.strip().upper()
    
    # 1. Search for standard time format with AM/PM (e.g., "10:00AM", "2:30 PM")
    # This regex looks for digits:digits followed optionally by space, then AM or PM
    match_ampm = re.search(r'(\d{1,2}:\d{2})\s*([AP]M)', time_str)
    if match_ampm:
        time_digits = match_ampm.group(1)
        meridiem = match_ampm.group(2)
        try:
            # Reconstruct clean string "10:00AM"
            t = datetime.strptime(f"{time_digits}{meridiem}", "%I:%M%p")
            return t.strftime("%I:%M %p")
        except ValueError:
            pass # Fall through to try other formats if this fails
            
    # 2. Search for 24h format or time without AM/PM (e.g., "14:00", "09:00")
    match_simple = re.search(r'(\d{1,2}:\d{2})', time_str)
    if match_simple:
        time_digits = match_simple.group(1)
        try:
            t = datetime.strptime(time_digits, "%H:%M")
            return t.strftime("%I:%M %p")
        except ValueError:
            pass

    return None

def parse_date(date_str):
    """Parses dates like '01/26/2026' or '26/01/2026'."""
    # Try MM/DD/YYYY first (common in US/software), then DD/MM/YYYY
    formats = ["%m/%d/%Y", "%d/%m/%Y", "%Y-%m-%d"]
    for fmt in formats:
        try:
            return datetime.strptime(date_str.strip(), fmt)
        except ValueError:
            continue
    return None

def map_day_to_weekday(day_str):
    """Maps day abbreviations to weekday number (Mon=0, Sun=6)."""
    d = day_str.lower()[:2] # Get first 2 chars
    mapping = {
        'mo': 0, 'm': 0,
        'tu': 1, 't': 1,
        'we': 2, 'w': 2,
        'th': 3, 'r': 3,
        'fr': 4, 'f': 4,
        'sa': 5, 's': 5,
        'su': 6
    }
    return mapping.get(d)

def generate_events(class_info):
    """
    Expands a class entry into individual events for every week in the date range.
    """
    events = []
    
    start_date_obj = parse_date(class_info['start_date_range'])
    end_date_obj = parse_date(class_info['end_date_range'])
    
    if not start_date_obj or not end_date_obj:
        print(f"Warning: Could not parse dates for {class_info['subject']}")
        return []

    # Identify target weekday
    target_weekday = map_day_to_weekday(class_info['day'])
    if target_weekday is None:
        return []

    # Iterate through every day from start to end date
    current_date = start_date_obj
    while current_date <= end_date_obj:
        if current_date.weekday() == target_weekday:
            # It's a match, create the event
            date_str = current_date.strftime("%m/%d/%Y") # Google Calendar prefers MM/DD/YYYY
            
            events.append({
                "Subject": class_info['subject'],
                "Start Date": date_str,
                "Start Time": parse_time(class_info['start_time']),
                "End Date": date_str,
                "End Time": parse_time(class_info['end_time']),
                "All Day Event": "False",
                "Description": class_info.get('description', ''),
                "Location": class_info.get('location', ''),
                "Private": "False"
            })
            
            # Jump 7 days
            current_date += timedelta(days=7)
        else:
            # Move to next day
            current_date += timedelta(days=1)
            
    return events

def extract_schedule_from_pdf(pdf_path):
    print(f"Opening {pdf_path}...")
    all_events = []
    
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages):
            print(f"Processing Page {page_num + 1}...")
            tables = page.extract_tables()
            
            for table in tables:
                # Iterate rows to find class data
                # We look for rows that likely contain date ranges and times
                # Regex for common date format MM/DD/YYYY or DD/MM/YYYY
                date_pattern = re.compile(r'\d{1,2}/\d{1,2}/\d{4}')
                time_pattern = re.compile(r'\d{1,2}:\d{2}')
                
                header_found = False
                
                for row in table:
                    # Clean None values
                    row = [cell.replace('\n', ' ') if cell else '' for cell in row]
                    row_text = " ".join(row)
                    
                    # Heuristic: Find rows that look like data
                    # Expecting: [Class Name] ... [Days] [Time] ... [Location] ... [Date - Date]
                    
                    dates = date_pattern.findall(row_text)
                    times = time_pattern.findall(row_text)
                    
                    if len(dates) >= 2 and len(times) >= 1:
                        try:
                            # 1. Find the dates (Start and End)
                            start_date_range = dates[0]
                            end_date_range = dates[1]
                            
                            # 2. Find the times
                            time_cell = next((c for c in row if time_pattern.search(c) and '-' in c), None)
                            if not time_cell: continue
                            
                            time_parts = time_cell.split('-')
                            start_time = time_parts[0].strip()
                            end_time = time_parts[1].strip()
                            
                            # 3. Find the Day (Mon, Tue, etc)
                            day_pattern = re.compile(r'\b(Mo|Tu|We|Th|Fr|Sa|Su|Mon|Tue|Wed|Thu|Fri|Sat|Sun)\b', re.IGNORECASE)
                            day_cell = next((c for c in row if day_pattern.search(c)), None)
                            if not day_cell: 
                                day_match = day_pattern.search(time_cell)
                                if day_match:
                                    day = day_match.group(0)
                                else:
                                    continue
                            else:
                                day = day_pattern.search(day_cell).group(0)
                                
                            # 4. Find Subject/Class Name
                            subject = row[0] + " " + row[1]
                            subject = subject.strip()
                            
                            # 5. Location
                            location = ""
                            for cell in row:
                                if "Bldg" in cell or re.search(r'\d\.\d{3}', cell):
                                    location = cell
                                    break
                            
                            class_info = {
                                'subject': subject,
                                'day': day,
                                'start_time': start_time,
                                'end_time': end_time,
                                'start_date_range': start_date_range,
                                'end_date_range': end_date_range,
                                'location': location,
                                'description': f"Extracted from {pdf_path}"
                            }
                            
                            new_events = generate_events(class_info)
                            all_events.extend(new_events)
                            print(f"  Found: {subject} on {day}")
                            
                        except Exception as e:
                            print(f"  Error parsing row: {row_text[:30]}... ({e})")
                            continue

    return all_events

def write_csv(events, output_filename=DEFAULT_OUTPUT_CSV):
    if not events:
        print("No events found! The PDF structure might be image-based or unsupported.")
        return

    print(f"Writing {len(events)} events to {output_filename}...")
    with open(output_filename, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.DictWriter(file, fieldnames=CSV_HEADERS)
        writer.writeheader()
        writer.writerows(events)
    print("Done!")

def launch_gui():
    """Launches the Tkinter GUI."""
    
    def browse_file():
        filename = filedialog.askopenfilename(
            title="Select Schedule PDF",
            filetypes=[("PDF Files", "*.pdf")]
        )
        if filename:
            entry_pdf.delete(0, tk.END)
            entry_pdf.insert(0, filename)

    def browse_output():
        filename = filedialog.asksaveasfilename(
            title="Save CSV As",
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv")],
            initialfile=DEFAULT_OUTPUT_CSV
        )
        if filename:
            entry_output.delete(0, tk.END)
            entry_output.insert(0, filename)

    def start_conversion():
        pdf_path = entry_pdf.get().strip()
        output_name = entry_output.get().strip()
        
        # Validation
        if not pdf_path:
            messagebox.showerror("Error", "Please select a PDF file first.")
            return
        
        if not os.path.exists(pdf_path):
            messagebox.showerror("Error", "The selected file does not exist.")
            return
            
        if not output_name:
            output_name = DEFAULT_OUTPUT_CSV
            
        if not output_name.lower().endswith('.csv'):
            output_name += '.csv'

        # Ensure directory exists if a path is provided
        output_dir = os.path.dirname(output_name)
        if output_dir and not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir)
            except OSError as e:
                messagebox.showerror("Error", f"Could not create output directory:\n{e}")
                return

        # Conversion Process
        btn_convert.config(state=tk.DISABLED, text="Processing...")
        status_label.config(text="Extracting schedule...", fg="blue")
        root.update()

        try:
            events = extract_schedule_from_pdf(pdf_path)
            
            if not events:
                status_label.config(text="No events found.", fg="red")
                messagebox.showwarning("Warning", "No events were found in the PDF.\nPlease check if the format matches.")
            else:
                write_csv(events, output_name)
                status_label.config(text="Success!", fg="green")
                messagebox.showinfo("Success", f"Converted {len(events)} events!\nSaved to: {output_name}")
                
        except Exception as e:
            status_label.config(text="Error occurred.", fg="red")
            messagebox.showerror("Error", f"An error occurred:\n{str(e)}")
            import traceback
            traceback.print_exc()
        finally:
            btn_convert.config(state=tk.NORMAL, text="Convert to CSV")

    # GUI Setup
    root = tk.Tk()
    root.title("Schedule to Calendar Converter")
    root.geometry("600x250") # Slightly wider for paths
    
    # Input Frame
    frame_input = tk.Frame(root, pady=10)
    frame_input.pack(fill=tk.X, padx=20)
    
    lbl_pdf = tk.Label(frame_input, text="PDF File:", font=("Arial", 10, "bold"))
    lbl_pdf.pack(anchor="w")
    
    entry_pdf = tk.Entry(frame_input, width=40)
    entry_pdf.pack(side=tk.LEFT, fill=tk.X, expand=True)
    
    btn_browse = tk.Button(frame_input, text="Browse...", command=browse_file)
    btn_browse.pack(side=tk.RIGHT, padx=(5, 0))

    # Output Frame
    frame_output = tk.Frame(root, pady=10)
    frame_output.pack(fill=tk.X, padx=20)
    
    lbl_out = tk.Label(frame_output, text="Output CSV Path:", font=("Arial", 10, "bold"))
    lbl_out.pack(anchor="w")
    
    entry_output = tk.Entry(frame_output, width=40)
    entry_output.insert(0, DEFAULT_OUTPUT_CSV)
    entry_output.pack(side=tk.LEFT, fill=tk.X, expand=True)

    btn_browse_out = tk.Button(frame_output, text="Browse...", command=browse_output)
    btn_browse_out.pack(side=tk.RIGHT, padx=(5, 0))

    # Action Frame
    frame_action = tk.Frame(root, pady=20)
    frame_action.pack(fill=tk.X, padx=20)
    
    btn_convert = tk.Button(frame_action, text="Convert to CSV", command=start_conversion, 
                            bg="#4CAF50", fg="white", font=("Arial", 11, "bold"), height=2)
    btn_convert.pack(fill=tk.X)
    
    status_label = tk.Label(root, text="Ready", fg="gray")
    status_label.pack()

    root.mainloop()

if __name__ == "__main__":
    # Check if arguments are passed. If so, use CLI. If not, use GUI.
    if len(sys.argv) > 1:
        parser = argparse.ArgumentParser(description="Convert Class Schedule PDF to Google Calendar CSV")
        parser.add_argument("input_pdf", nargs="?", help="Path to the PDF file")
        parser.add_argument("--output", "-o", default=DEFAULT_OUTPUT_CSV, help="Output CSV filename")
        
        args = parser.parse_args()
        
        input_pdf = args.input_pdf
        
        # CLI fallback logic for file selection
        if not input_pdf:
            pdfs = [f for f in os.listdir('.') if f.lower().endswith('.pdf')]
            if len(pdfs) == 1:
                input_pdf = pdfs[0]
            elif len(pdfs) > 1:
                print("Multiple PDFs found. Please specify one or use the GUI.")
                sys.exit(1)
        
        if input_pdf and os.path.exists(input_pdf):
            try:
                events = extract_schedule_from_pdf(input_pdf)
                write_csv(events, args.output)
            except Exception as e:
                print(f"Error: {e}")
        else:
            print("File not found. Try running without arguments to use the GUI.")
    else:
        launch_gui()