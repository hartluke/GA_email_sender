# CSM General Accounting Email Sender

This project was made for the General Accounting team at Colorado School of Mines MAPS and is designed to automate the process of sending emails with data extracted from various report spreadsheets. The emails are formatted using HTML templates and the data is inserted into these templates based on a configuration file.

## Features

- Extracts data from Excel files
- Formats emails using HTML templates
- Sends emails using the Outlook application and Pywin library
- Supports multiple email templates
- Configuration file allows for flexible mapping of Excel data to email variables
- Exported to an executable file with an easy-to-use interface

## Usage

1. Set your default Outlook account to the account you wish to send the emails from
2. Run the .exe file in the dist folder
3. Input the excel file from the report spreadsheet

## Requirements

- Windows 10+
- Outlook Desktop Application
- CSM Global Protect VPN

## Note

This script is designed to work with the Outlook application on a Windows computer. Please ensure that Outlook is installed and configured on your machine before running the script. This app **will not** work on any OS other than Windows.

Regenerate executable: python -m PyInstaller --onefile --name GA_email_sender --distpath ./dist --workpath ./build --specpath ./spec app.py --noconsole --icon="//files.mines.edu/shared/Accounting/Accounting (Shared)/1Procedure Library - Office of Controller/Luke Projects/GA_email_sender/src/assets/logo.ico"