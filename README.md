# SUTD-Calander-Import
PDF to CSV converter for easy import to google calander
Prerequisites

Python 3.6 or higher

pip (Python package installer)

Installation

Clone or Download this repository.

Install dependencies:
The only external library required is pdfplumber.

pip install pdfplumber


How to Get Your Schedule PDF 

Log in to MyPortal and navigate to My Weekly Schedule.

Select List View.

Scroll all the way down to the bottom of the page.

Select Printer Friendly Page.

Press Ctrl + P (or Command + P on Mac) to open the print dialog.

Choose Save as PDF as the destination and save it to a directory of your choice.

Usage

Method 1: Graphical User Interface (GUI)

This is the easiest way to use the tool.

Run the script without any arguments:

python convert_schedule.py



A window will appear.

Click "Browse..." to select your Schedule PDF.

(Optional) Click "Browse..." to choose where to save the CSV file.

Click "Convert to CSV".

Method 2: Command Line (CLI)

You can specify files directly in the terminal, which is useful for scripts or quick conversions.

Syntax:

python convert_schedule.py [input_pdf_path] --output [output_csv_path]



Importing to Google Calendar

Once you have generated the CSV file:

Open Google Calendar on your computer.

Click the Settings gear icon (top right) > Settings.

In the left sidebar, click Import & export.

Under "Import", click Select file from your computer and choose your generated CSV file.

Select the specific calendar you want to add these events to (e.g., your primary calendar).

Click Import.

Troubleshooting

"No events found": The PDF might be an image scan (not selectable text) or have a radically different table layout. This script relies on pdfplumber finding text tables.

Wrong Times: Check the CSV file. If times are missing, the PDF might have unusual formatting (e.g., newlines splitting the time). The script attempts to clean "Mon 10:00AM" into "10:00 AM".

