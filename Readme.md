# App Script RBS Support Automation

This project automates the processing of support emails for RBS using Google Apps Script. It extracts relevant portions of email bodies, organizes them into Google Sheets, and helps streamline the support process.

## Features
- Extracts relevant portions of email bodies.
- Organizes data into Google Sheets by month and year.
- Automatically creates new sheets and sets up headers if they don't exist.

## Setup

### Clone the repository

1. **Open the project in Google Apps Script:**
   - Go to [Google Apps Script](https://script.google.com/).
   - Create a new project and copy the contents of the `Code.js` file into the script editor.

2. **Set up Google Sheets:**
   - Create a new Google Sheet.
   - Share the sheet with the email address of the Google Apps Script service account.

3. **Configure script properties:**
   - In the Google Apps Script editor, go to `File > Project properties > Script properties`.
   - Add the following properties:
     - `SPREADSHEET_ID`: The ID of the Google Sheet you created.

## Usage

### Run the script
- In the Google Apps Script editor, click the run button to execute the script.
- The script will process the emails and update the Google Sheet accordingly.

### Automate the script
- Set up a time-driven trigger in Google Apps Script to run the script periodically:
  - Go to `Edit > Current project's triggers` and add a new trigger.

## Contributing

1. **Fork the repository:**
   - Click the "Fork" button at the top right of the repository page.

2. **Clone the repository:**
   - Create a new branch.

3. **Make your changes:**
   - Open the `Code.js` file in your preferred code editor (e.g., Visual Studio Code).
   - Make the necessary changes.
   # Google Apps Script RBS Support Automation

This project automates the processing of support emails for RBS using Google Apps Script. It extracts relevant portions of email bodies, organizes them into Google Sheets, and helps streamline the support process.

## Features
- Extracts relevant portions of email bodies.
- Organizes data into Google Sheets by month and year.
- Automatically creates new sheets and sets up headers if they don't exist.

## Setup

### Clone and Push Using `clasp`

#### 1. Install `clasp` Globally

Install `clasp` using npm:

```bash
npm install -g @google/clasp

4. **Commit your changes:**
   - Commit the changes to your branch.

5. **Push to the branch:**
   - Push your changes to the remote repository.

6. **Create a pull request:**
   - Go to the repository page on GitHub and click the "New pull request" button.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.