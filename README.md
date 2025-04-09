Document Automation with Excel Data
This project automates the process of generating a Word document (.docx) using data from an Excel file (.xlsx). It processes the data and fills in placeholders within a Word template file, replacing them with information from the Excel rows. The output document contains the data from each Excel row in a formatted manner.

Features
Excel to Word Automation: Automatically fills in placeholders in a Word document template with values from an Excel file.

Star Rating Conversion: Converts numeric ratings (with decimal values like .0 and .5) into a star format (★ for full stars, ⯪ for half stars, and ☆ for empty stars).

Pagination: Automatically inserts page breaks between data entries to ensure each entry is placed on a new page.

Prerequisites
To use this project, ensure you have the following installed:

Python 3.x

Required Python libraries:

pandas: Used for reading and manipulating Excel files.

python-docx: Used for working with Word documents.

You can install the required libraries by running:


pip install pandas python-docx

Excel File (Sample.xlsx): Your data should be stored in an Excel file with the following columns:

Roll Number
Name
Email ID
Citi Score
Rating

Template Word Document (template.docx): A Word document template with placeholders like {{Roll Number}}, {{Name}}, {{Email ID}}, {{Citi Score}}, and {{Rating}}. These placeholders will be replaced by data from the Excel file.

Setup Instructions
Clone the repository to your local machine:

git clone https://github.com/preetamkumar017/excel-to-word-automation.git

Place your Sample.xlsx Excel file and template.docx Word template in the same directory as the Python script (script.py).

Ensure your Excel file has the necessary columns:

Roll Number: The student's roll number.
Name: The student's name.
Email ID: The student's email address.
Citi Score: The student's Citi Score (numeric, preferably with .0 or .5 values).
Rating: The rating of the student, which will be converted to stars.

Usage
Once the prerequisites are set up, run the Python script:

python script.py
This will:

Read data from Sample.xlsx
Replace placeholders in template.docx with values from the Excel file

Add page breaks after each row
Save the generated document as output.docx in the same directory



License
This project is licensed under the MIT License - see the LICENSE.md file for details.
