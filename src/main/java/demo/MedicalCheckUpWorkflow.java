package demo;

import helpers.DateHelper;
import helpers.FormHelper;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;

import org.apache.poi.ss.usermodel.*;

import java.time.LocalDate;


import static demo.WorkflowSupport.*;

final class MedicalCheckUpWorkflow {
    static void addMedCheck(WebDriver driver, Sheet medCheckSheet, int rowIndex) {
        int spreadsheetRow = rowIndex + 1;
        String catName = getCellValueAsString(getCell(medCheckSheet, rowIndex, 0)).trim();
        String neededOrGiven = getCellValueAsString(getCell(medCheckSheet, rowIndex, 1)).trim();
        LocalDate examDate = DateHelper.getCellValueAsDate(getCell(medCheckSheet, rowIndex, 2));
        String vet = getCellValueAsString(getCell(medCheckSheet, rowIndex, 3));
        String vetClinic = getCellValueAsString(getCell(medCheckSheet, rowIndex, 4));
        String genComments = getCellValueAsString(getCell(medCheckSheet, rowIndex, 5));
        String hidComments = getCellValueAsString(getCell(medCheckSheet, rowIndex, 6));
        String recommendations = getCellValueAsString(getCell(medCheckSheet, rowIndex, 7));
        String weight = getCellValueAsString(getCell(medCheckSheet, rowIndex, 8));

        if (!"Needed".equalsIgnoreCase(neededOrGiven) && !"Given".equalsIgnoreCase(neededOrGiven)) {
            throw new IllegalArgumentException(
                    "Medical Check-Up!B" + spreadsheetRow + " must be either 'Needed' or 'Given'.");
        }
        if (examDate == null) {
            throw new IllegalArgumentException(
                    "Medical Check-Up!C" + spreadsheetRow + " must contain a valid date.");
        }

        openAnimalFromSearch(driver, catName, "Medical Check-Up!A" + spreadsheetRow);
        openAnimalHistoryAddForm(
                driver, catName, "Medical History", "med.php", "Add a Medical Check-Up").click();

        String dateFieldName = "Needed".equalsIgnoreCase(neededOrGiven) ? "nd" : "gd";
        DateHelper.enterDateText(driver, dateFieldName, examDate, true);
        FormHelper.enterText(driver, "v", vet);
        FormHelper.enterText(driver, "vc", vetClinic);
        FormHelper.enterText(driver, "gc", genComments);
        FormHelper.enterText(driver, "hc", hidComments);
        FormHelper.enterText(driver, "r", recommendations);
        FormHelper.enterText(driver, "w", weight);
        By submitButton = By.cssSelector("input[value='Add']");
        String formUrl = driver.getCurrentUrl();
        driver.findElement(submitButton).click();
        waitForFormSubmission(driver, submitButton, formUrl, "Medical Check-Up", spreadsheetRow);
    }
}
