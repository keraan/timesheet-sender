# Timesheet Automation Tool for Masman

The **Timesheet Automation Tool** is a Python script that automates the process of creating and sending timesheets via Gmail. It utilizes Google APIs for authentication, email sending, and Google Sheets integration to create a comprehensive timesheet report and deliver it to the specified recipient. This tool is particularly useful for individuals who need to submit timesheets regularly and want to streamline the process.

## Features

- Automatically generates a timesheet report for the past week (last 7 days) based on the user's provided schedule and working hours.
- Uses Google APIs to authenticate and send emails, ensuring secure and reliable delivery of the timesheet report.
- Integrates with Google Sheets to load and edit the timesheet, saving the modified report for future reference.
- Supports input of start and end times for each working day, calculating the total hours worked and appending the information to the timesheet.
- Customizable configurations through the `config.json` file, allowing users to define personal details, recipient's email, and file paths.

## Getting Started

Follow these steps to set up and use the Timesheet Automation Tool:

1. **Clone the Repository**: Clone this repository to your local machine.

2. **Install Dependencies**: Install the required Python packages by running the following command:

```
pip install -r requirements.txt
```
   
3. **Configure the Tool**: Edit the `config.json` file with your personal details, including your first name, last name, email, and recipient's email. You'll also need to provide the paths to your `client_secret_file.json` (for Google API authentication) and the `original_file` (the template timesheet file you want to use).

4. **Run the Script**: Run the `mailSender.py` script using the following command:

```
python3 mailSender.py
```

Follow the prompts to input the working days, start times, and end times for each day.

5. **Email Delivery**: The tool will automatically create a timesheet report in Excel format and send it as an attachment via Gmail to the specified recipient.

6. **File Saving**: A copy of the timesheet report will be saved locally for your records.

## Note

- The script uses Google APIs for authentication. Ensure you have the necessary credentials (`client_secret_file.json`) and permissions to send emails using your Gmail account.

- The script assumes a Monday to Saturday working week. If your schedule is different, you can modify the `input_day` function to suit your needs.

- The timesheet report is generated in Excel format and is attached to the email.

- Please ensure you have a stable internet connection while running the script, as it relies on Google services for authentication and email sending.

- Before using the tool, familiarize yourself with the Gmail API and Google Sheets API documentation.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

**Disclaimer**: This tool is provided as-is and without warranty. Use it responsibly and ensure you have the necessary permissions to access and modify the provided files.
