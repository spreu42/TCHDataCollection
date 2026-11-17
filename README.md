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
program yet. The demo and production launchers read separate workbook copies
from `src/main/resources` and must be started using the developer instructions
below.

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

### Run the demo program

Run `demo.FurstTestDM.main` from an IDE, with the working directory set to the
project root. The program reads `src/main/resources/DCWorkbookDM.xlsm` and signs
into the iShelters demo/test site at
`https://ishelter.demoshelters.com/login.php`.

### Run the production program

Run `prod.FurstTest.main` from an IDE, with the working directory set to the
project root. It reads `src/main/resources/DCWorkbook.xlsm` and signs into
production at `https://ishelter.ishelters.com/`.

Demo workflows are under `src/main/java/demo/` and end in `DM`; production
workflows with their original class names are under `src/main/java/prod/`.
Both use the shared utilities under `src/main/java/helpers/`, including
`WorkflowSupport`.

Each populated workflow sheet is processed independently, one data row at a
time. New Animal, Check-In (including the optional Littermates? workflow),
Medical Procedure, Vaccinations, Medical Check-Up, and Prescriptions rows are
processed in spreadsheet order. For a litter row, column Q must be `Yes`, `No`,
or blank. When it is `Yes` and column R identifies the mother, the program
starts the litter entry from her record. If she is also listed in column B, she
is omitted from the litter information; otherwise, all cats in column B are
added as her kittens. Column R may be blank for a litter without a mother; in
that case, the program starts from the first cat in column B and adds all listed
cats to the litter. The litter record is submitted before the row's Check-In
record. Consecutive
prescriptions for the same animal reuse its current history page. After a row's
submission is confirmed, that row is deleted and the workbook is saved
immediately, so a later failure does not leave already-completed rows queued
for resubmission. Rows that fail validation or submission remain in the
workbook. Rows whose animal-name cell contains the template placeholder
`All Fields` are skipped with a diagnostic message.
The Litter Information and Add a Litter controls are each awaited for up to 10
seconds, checking once per second. The Litter History panel is embedded in the
animal profile, so the workflow locates the Add a Litter link in its frame.
The Check-In Move Requested field accepts `Y` or `Yes`, case-insensitively.
On the New Intake - Add sheet, the first cat uses the `Add Animal` link and
subsequent cats use the `Add Another Animal` image link.
This selection is based on rows processed during the current run; after a
restart, the first remaining row is treated as the initial cat.
When a New Intake - Add row reaches the Similar Animals page, the error output
includes the page title and URL and saves a screenshot under the system
temporary directory in `tch-workflow-failures`; the exact screenshot path is
printed to the console. If the `Add a Litter` control is missing after opening
Litter Information, the workflow reports the page details and saves a
screenshot in the same directory. Check-In selections are made from the site's
autocomplete suggestions; rejected submissions include the exact alert text
and save a screenshot for review.

Each launcher writes to its own workbook: the demo program uses
`src/main/resources/DCWorkbookDM.xlsm`, and production uses
`src/main/resources/DCWorkbook.xlsm`. Neither reads or updates the Desktop
workbook. Close the selected workbook in Excel before running, because
completed rows must be saved back to the file.

## Future investigation

- Investigate using AI OCR to read medical intake forms directly from fax.
