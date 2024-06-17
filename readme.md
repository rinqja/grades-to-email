# Student Grades Emailer

This Python script reads student grades from an Excel file and sends an email to a specified recipient with the formatted grades.

## Features

- Reads grades from an Excel file.
- Handles missing grades by filling them with 0.
- Formats grades as integers in the email body.
- Sends an email with the formatted grades to a specified recipient.

## Prerequisites

- Python 3.x
- pandas library
- openpyxl library
- smtplib library

## Installation

1. Clone the repository:

   ```bash
   git clone https://github.com/yourusername/student-grades-emailer.git
   cd student-grades-emailer
   ```

2. Install the required libraries:
   ```bash
   pip install pandas openpyxl
   ```

## Configuration

1. Create a `config.py` file with the following content and fill in your details:

   ```python
   FROM_EMAIL = 'your_email@gmail.com'
   PASSWORD = 'your_app_specific_password'
   TO_EMAIL = 'recipient_email@example.com'
   ```

2. Set the path to your Excel file in the `main` function:
   ```python
   file_path = 'path/to/your/grades.xlsx'
   ```

## Usage

Run the script:

```bash
python grades.py
```
