# TCH Data Collection

This tool transfers animal and medical information from the DCWorkbook
spreadsheet to iShelters.

## User instructions

These are the steps the finished, everyday-use setup is intended to have:

1. Open the **DCWorkbook** spreadsheet on the Desktop.
2. Fill in the information for the animal(s). Check your entries before running
   the tool.
3. Click the **TCH Data Collection** button on the Desktop to send the
   information to iShelters.
4. Wait for the process to finish before closing the browser.

**Important:** The Desktop spreadsheet and run button are not connected to the
program yet. For now, the program reads the copy of `DCWorkbook.xlsm` in the
project's `src/main/resources` folder and must be started using the developer
instructions below.

## Tech setup and current run instructions

### Requirements

- Java
- Maven
- Google Chrome

### Set up credentials

The program requires the iShelters username and password in local environment
variables. It does not read credentials from the spreadsheet.

In Command Prompt, set the variables in the same window that will run the
program:

```cmd
set TCH_USER=your-username
set TCH_PASSWORD=your-password
```

Keep credentials private and do not commit them to the repository. The program
stops with an error if either variable is missing or blank.

### Run the current program

Run `FurstTest.main` from an IDE, with the working directory set to the project
root. The program reads `src/main/resources/DCWorkbook.xlsm` and signs into the
iShelters demo/test site at `https://ishelter.demoshelters.com/login.php`. The
production URL is commented out in the source.

The program currently processes the first data row in each populated sheet.
Workbook row clearing and saving are disabled for repeatable testing, so running
it again submits the same data again. Use only appropriate test records.

## Future investigation

- Investigate using AI OCR to read medical intake forms directly from fax.
