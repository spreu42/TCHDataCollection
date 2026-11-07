package demo;

import helpers.DateHelper;
import helpers.FormHelper;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;

import org.apache.poi.ss.usermodel.*;

import java.time.LocalDate;


import static demo.WorkflowSupport.*;

final class VaccinationWorkflow {
    static void addVaccinations(WebDriver driver, Sheet vaccSheet, int rowIndex) {
        int spreadsheetRow = rowIndex + 1;
        String catName = getCellValueAsString(getCell(vaccSheet, rowIndex, 0)).trim();
        String neededOrGiven = getCellValueAsString(getCell(vaccSheet, rowIndex, 1)).trim();
        LocalDate vaccDate = DateHelper.getCellValueAsDate(getCell(vaccSheet, rowIndex, 2));
        String vaccType = getCellValueAsString(getCell(vaccSheet, rowIndex, 3)).trim();
        String vet = getCellValueAsString(getCell(vaccSheet, rowIndex, 4));
        String vetClinic = getCellValueAsString(getCell(vaccSheet, rowIndex, 5));
        String genComments = getCellValueAsString(getCell(vaccSheet, rowIndex, 6));
        String hidComments = getCellValueAsString(getCell(vaccSheet, rowIndex, 7));

        if (!"Needed".equalsIgnoreCase(neededOrGiven) && !"Given".equalsIgnoreCase(neededOrGiven)) {
            throw new IllegalArgumentException(
                    "Vaccinations!B" + spreadsheetRow + " must be either 'Needed' or 'Given'.");
        }
        if (vaccDate == null) {
            throw new IllegalArgumentException(
                    "Vaccinations!C" + spreadsheetRow + " must contain a valid date.");
        }

        String vaccineCode;
        if ("Feline Panleukopenia".equalsIgnoreCase(vaccType)) {
            vaccineCode = "26569";
        } else if ("Rabies".equalsIgnoreCase(vaccType)) {
            vaccineCode = "26574";
        } else if ("Rhino Trachitis/Feline Calicivirus".equalsIgnoreCase(vaccType)) {
            vaccineCode = "26571";
        } else {
            throw new IllegalArgumentException(
                    "Vaccinations!D" + spreadsheetRow + " contains an unsupported vaccine: '"
                            + vaccType + "'.");
        }

        openAnimalFromSearch(driver, catName, "Vaccinations!A" + spreadsheetRow);
        openAnimalHistoryAddForm(
                driver, catName, "Vaccination History", "vaccine.php", "Add a Vaccination Record").click();

        String dateFieldName = "Needed".equalsIgnoreCase(neededOrGiven) ? "nd" : "gd";
        DateHelper.enterDateText(driver, dateFieldName, vaccDate, true);

        driver.findElement(By.cssSelector("input[value='" + vaccineCode + "']")).click();
        FormHelper.enterText(driver, "v", vet);
        FormHelper.enterText(driver, "vc", vetClinic);
        FormHelper.enterText(driver, "gc", genComments);
        FormHelper.enterText(driver, "hc", hidComments);
        By submitButton = By.cssSelector("input[value='Add']");
        String formUrl = driver.getCurrentUrl();
        driver.findElement(submitButton).click();
        waitForFormSubmission(driver, submitButton, formUrl, "Vaccinations", spreadsheetRow);
    }
}
