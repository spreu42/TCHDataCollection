package demo;

import helpers.DateHelper;
import helpers.FormHelper;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;

import org.apache.poi.ss.usermodel.*;

import java.time.LocalDate;


import static demo.WorkflowSupport.*;

final class MedicalProcedureWorkflow {
    static void addProcedure(WebDriver driver, Sheet procedureSheet, int rowIndex) {
        int spreadsheetRow = rowIndex + 1;
        String catName = getCellValueAsString(getCell(procedureSheet, rowIndex, 0)).trim();
        String neededOrPerformed = getCellValueAsString(getCell(procedureSheet, rowIndex, 1)).trim();
        LocalDate procedureDate = DateHelper.getCellValueAsDate(getCell(procedureSheet, rowIndex, 2));
        String procedureType = getCellValueAsString(getCell(procedureSheet, rowIndex, 3)).trim();
        String vet = getCellValueAsString(getCell(procedureSheet, rowIndex, 4)).trim();
        String vetClinic = getCellValueAsString(getCell(procedureSheet, rowIndex, 5)).trim();
        String genComments = getCellValueAsString(getCell(procedureSheet, rowIndex, 6));
        String hidComments = getCellValueAsString(getCell(procedureSheet, rowIndex, 7));

        if (!"Needed".equalsIgnoreCase(neededOrPerformed)
                && !"Performed".equalsIgnoreCase(neededOrPerformed)) {
            throw new IllegalArgumentException(
                    "Medical Procedure!B" + spreadsheetRow + " must be either 'Needed' or 'Performed'.");
        }
        if (procedureDate == null) {
            throw new IllegalArgumentException(
                    "Medical Procedure!C" + spreadsheetRow + " must contain a valid date.");
        }
        if (procedureType.isEmpty()) {
            throw new IllegalArgumentException(
                    "Medical Procedure!D" + spreadsheetRow + " must contain a procedure type.");
        }

        openAnimalFromSearch(driver, catName, "Medical Procedure!A" + spreadsheetRow);
        openAnimalHistoryAddForm(
                driver, catName, "Medical History", "med.php", "Add a Medical Procedure").click();

        String dateFieldName = "Needed".equalsIgnoreCase(neededOrPerformed) ? "nd" : "gd";
        DateHelper.enterDateText(driver, dateFieldName, procedureDate, true);

        FormHelper.selectByVisibleText(driver, "tp", procedureType);
        if (!vet.isEmpty()) {
            FormHelper.selectByVisibleText(driver, "v", vet);
        }
        if (!vetClinic.isEmpty()) {
            FormHelper.selectByVisibleText(driver, "vc", vetClinic);
        }
        FormHelper.enterText(driver, "gc", genComments);
        FormHelper.enterText(driver, "hc", hidComments);
        By submitButton = By.cssSelector("input#submit[type='submit'][value='Add']");
        String formUrl = driver.getCurrentUrl();
        driver.findElement(submitButton).click();
        waitForFormSubmission(driver, submitButton, formUrl, "Medical Procedure", spreadsheetRow);
    }
}
