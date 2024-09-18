## Mailbill - A simple and efficient tool for the extraction of electronic invoices
Mailbill is a tool that allows you to extract .xml files from your Gmail, process them and save them in a spreadsheet document (.xls).
### Note: This tool is intended to be used with Ecuador's electronic invoicing format.
## Usage:
- Install the required packages using the following command: `pip install -r requirements.txt`
- Create an `.env` file where you must enter your email (`EMAIL_ACCOUNT`), [password](https://knowledge.workspace.google.com/kb/how-to-create-app-passwords-000009237?hl=es-419) (`PASSWORD`) and the path (`XLS_FILE`) where your .xls file is located (if it does not exist, it will be created automatically).
- Run the command `python3 main.py` and leave it running for as long as you want to automatically extract the xml invoices.
