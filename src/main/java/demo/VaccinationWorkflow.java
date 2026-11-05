package demo;

import helpers.DateHelper;
import helpers.FormHelper;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;

import org.apache.poi.ss.usermodel.*;

import java.time.LocalDate;


import static demo.WorkflowSupport.*;

final class VaccinationWorkflow {
    static void addVaccinations(WebDriver driver, Sheet vaccSheet) {
        if (!hasData(vaccSheet)) {
            return;
        }

        String catName = getCellValueAsString(getCell(vaccSheet, 1, 0)).trim();
        String neededOrGiven = getCellValueAsString(getCell(vaccSheet, 1, 1)).trim();
        LocalDate vaccDate = DateHelper.getCellValueAsDate(getCell(vaccSheet, 1, 2));
        String vaccType = getCellValueAsString(getCell(vaccSheet, 1, 3)).trim();
        String vet = getCellValueAsString(getCell(vaccSheet, 1, 4));
        String vetClinic = getCellValueAsString(getCell(vaccSheet, 1, 5));
        String genComments = getCellValueAsString(getCell(vaccSheet, 1, 6));
        String hidComments = getCellValueAsString(getCell(vaccSheet, 1, 7));

        if (!"Needed".equalsIgnoreCase(neededOrGiven) && !"Given".equalsIgnoreCase(neededOrGiven)) {
            throw new IllegalArgumentException(
                    "Vaccinations!B2 must be either 'Needed' or 'Given'.");
        }
        if (vaccDate == null) {
            throw new IllegalArgumentException("Vaccinations!C2 must contain a valid date.");
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
                    "Vaccinations!D2 contains an unsupported vaccine: '" + vaccType + "'.");
        }

        openAnimalFromSearch(driver, catName, "Vaccinations!A2");
        openAnimalHistoryAddForm(
                driver, catName, "Vaccination History", "vaccine.php", "Add a Vaccination Record").click();

        String dateFieldName = "Needed".equalsIgnoreCase(neededOrGiven) ? "nd" : "gd";
        DateHelper.enterDateText(driver, dateFieldName, vaccDate, true);

        driver.findElement(By.cssSelector("input[value='" + vaccineCode + "']")).click();
        FormHelper.enterText(driver, "v", vet);
        FormHelper.enterText(driver, "vc", vetClinic);
        FormHelper.enterText(driver, "gc", genComments);
        FormHelper.enterText(driver, "hc", hidComments);
        driver.findElement(By.cssSelector("input[value='Add']")).click();

        // Row clearing is disabled so the same test data can be reused.
        // deleteRow(vaccSheet, 1);
    }
}

