package demo;

import helpers.DateHelper;
import helpers.FormHelper;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;

import org.apache.poi.ss.usermodel.*;

import java.time.LocalDate;


import static demo.WorkflowSupport.*;

final class MedicalProcedureWorkflow {
    static void addProcedure(WebDriver driver, Sheet procedureSheet) {
        if (!hasData(procedureSheet)) {
            return;
        }

        String catName = getCellValueAsString(getCell(procedureSheet, 1, 0)).trim();
        String neededOrPerformed = getCellValueAsString(getCell(procedureSheet, 1, 1)).trim();
        LocalDate procedureDate = DateHelper.getCellValueAsDate(getCell(procedureSheet, 1, 2));
        String procedureType = getCellValueAsString(getCell(procedureSheet, 1, 3)).trim();
        String vet = getCellValueAsString(getCell(procedureSheet, 1, 4)).trim();
        String vetClinic = getCellValueAsString(getCell(procedureSheet, 1, 5)).trim();
        String genComments = getCellValueAsString(getCell(procedureSheet, 1, 6));
        String hidComments = getCellValueAsString(getCell(procedureSheet, 1, 7));

        if (!"Needed".equalsIgnoreCase(neededOrPerformed)
                && !"Performed".equalsIgnoreCase(neededOrPerformed)
                && !"Given".equalsIgnoreCase(neededOrPerformed)) {
            throw new IllegalArgumentException(
                    "Medical Procedure!B2 must be either 'Needed' or 'Performed'.");
        }
        if (procedureDate == null) {
            throw new IllegalArgumentException("Medical Procedure!C2 must contain a valid date.");
        }
        if (procedureType.isEmpty()) {
            throw new IllegalArgumentException("Medical Procedure!D2 must contain a procedure type.");
        }

        openAnimalFromSearch(driver, catName, "Medical Procedure!A2");
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
        driver.findElement(By.cssSelector("input#submit[type='submit'][value='Add']")).click();

        // Row clearing is disabled so the same test data can be reused.
        // deleteRow(procedureSheet, 1);
    }
}

