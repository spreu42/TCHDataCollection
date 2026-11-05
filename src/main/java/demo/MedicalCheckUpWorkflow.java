package demo;

import helpers.DateHelper;
import helpers.FormHelper;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;

import org.apache.poi.ss.usermodel.*;

import java.time.LocalDate;


import static demo.WorkflowSupport.*;

final class MedicalCheckUpWorkflow {
    static void addMedCheck(WebDriver driver, Sheet medCheckSheet) {
        if (!hasData(medCheckSheet)) {
            return;
        }

        String catName = getCellValueAsString(getCell(medCheckSheet, 1, 0)).trim();
        String neededOrGiven = getCellValueAsString(getCell(medCheckSheet, 1, 1)).trim();
        LocalDate examDate = DateHelper.getCellValueAsDate(getCell(medCheckSheet, 1, 2));
        String vet = getCellValueAsString(getCell(medCheckSheet, 1, 3));
        String vetClinic = getCellValueAsString(getCell(medCheckSheet, 1, 4));
        String genComments = getCellValueAsString(getCell(medCheckSheet, 1, 5));
        String hidComments = getCellValueAsString(getCell(medCheckSheet, 1, 6));
        String recommendations = getCellValueAsString(getCell(medCheckSheet, 1, 7));
        String weight = getCellValueAsString(getCell(medCheckSheet, 1, 8));

        if (!"Needed".equalsIgnoreCase(neededOrGiven) && !"Given".equalsIgnoreCase(neededOrGiven)) {
            throw new IllegalArgumentException(
                    "Medical Check-Up!B2 must be either 'Needed' or 'Given'.");
        }
        if (examDate == null) {
            throw new IllegalArgumentException("Medical Check-Up!C2 must contain a valid date.");
        }

        openAnimalFromSearch(driver, catName, "Medical Check-Up!A2");
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
        driver.findElement(By.cssSelector("input[value='Add']")).click();

        // Row clearing is disabled so the same test data can be reused.
        // deleteRow(medCheckSheet, 1);
    }
}

