import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import config

# Function to read Excel file and extract relevant data
def read_excel(file_path):
    try:
        df = pd.read_excel(file_path)
        # Strip any leading/trailing spaces from column names and normalize them
        df.columns = df.columns.str.strip().str.replace('  ', ' ')
        df = df.fillna(0)  # Fill any blank cells with 0
        print("Column names:", df.columns)
        return df
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None

# Function to format the extracted data for email
def format_data_for_email(data):
    try:
        email_body = "Student Grades:\n\n"
        for index, row in data.iterrows():
            grades = ', '.join([f"Sprint {i+1}: {row[f'Sprint  {i+1}']}" for i in range(5)])   
            email_body += f"Name: {row['Student Grades']}, Grades: {grades}\n"
        return email_body
    except Exception as e:
        print(f"Error formatting email body: {e}")
        return None

# Function to send an email with the formatted data
def send_email(subject, body, to_email, from_email, password):
    try:
        msg = MIMEMultipart()
        msg['From'] = from_email
        msg['To'] = to_email
        msg['Subject'] = subject
        
        msg.attach(MIMEText(body, 'plain'))
        
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(from_email, password)
        text = msg.as_string()
        server.sendmail(from_email, to_email, text)
        server.quit()
        print("Email sent successfully!")
    except Exception as e:
        print(f"Error sending email: {e}")

# Main function
def main():
    file_path = 'C:/Users/rionk/Downloads/grades.xlsx'  # Update this path to your actual file path
    to_email = config.TO_EMAIL
    from_email = config.FROM_EMAIL
    password = config.PASSWORD

    data = read_excel(file_path)
    if data is not None:
        email_body = format_data_for_email(data)
        if email_body is not None:
            send_email('Student Grades', email_body, to_email, from_email, password)

if __name__ == "__main__":
    main()
