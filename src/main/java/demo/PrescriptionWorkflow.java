package demo;

import helpers.DateHelper;
import helpers.FormHelper;
import org.openqa.selenium.By;
import org.openqa.selenium.TimeoutException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import org.apache.poi.ss.usermodel.*;

import java.time.Duration;
import java.time.LocalDate;
import java.util.HashMap;
import java.util.Locale;
import java.util.Map;
import java.util.stream.Collectors;


import static demo.WorkflowSupport.*;

final class PrescriptionWorkflow {
    static void addPrescriptions(WebDriver driver, Sheet prescriptionSheet) {
        String currentAnimalName = null;
        Map<String, String> animalProfileUrls = new HashMap<>();
        for (int rowIndex = 1; rowIndex <= prescriptionSheet.getLastRowNum(); rowIndex++) {
            String catName = getCellValueAsString(getCell(prescriptionSheet, rowIndex, 0)).trim();
            if (catName.isEmpty() || "All Fields".equalsIgnoreCase(catName)) {
                continue;
            }
            if (currentAnimalName == null || !currentAnimalName.equalsIgnoreCase(catName)) {
                String animalKey = catName.toLowerCase(Locale.ROOT);
                String animalProfileUrl = animalProfileUrls.get(animalKey);
                if (animalProfileUrl == null) {
                    openAnimalFromSearch(driver, catName, "Prescriptions!A" + (rowIndex + 1));
                    animalProfileUrl = driver.getCurrentUrl();
                    animalProfileUrls.put(animalKey, animalProfileUrl);
                } else {
                    driver.navigate().to(animalProfileUrl);
                    new WebDriverWait(driver, Duration.ofSeconds(10))
                            .until(ExpectedConditions.presenceOfElementLocated(By.cssSelector("select.history")));
                }
                openAnimalHistoryAddForm(
                        driver, catName, "Medication Prescriptions", "medication.php",
                        "Add a Medication Prescription").click();
                currentAnimalName = catName;
            } else {
                WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
                By animalMedicationHistoryLink = By.linkText("Animal's Medication Prescriptions");
                By addMedicationLink = By.cssSelector("[title='Add a Medication Prescription']");
                try {
                    wait.until(ExpectedConditions.elementToBeClickable(animalMedicationHistoryLink)).click();
                    wait.until(ExpectedConditions.urlContains("medication.php?id="));
                    wait.until(ExpectedConditions.elementToBeClickable(addMedicationLink)).click();
                } catch (TimeoutException e) {
                    throw new IllegalStateException(
                            "Could not return to Medication Prescriptions for '" + catName
                                    + "' after submitting row " + rowIndex + ". Current URL: "
                                    + driver.getCurrentUrl(),
                            e);
                }
            }

            addPrescriptionRecord(driver, prescriptionSheet, rowIndex, catName);
            WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
            String responseUrl;
            try {
                responseUrl = wait.until(currentDriver -> {
                    String url = currentDriver.getCurrentUrl();
                    return url.contains("medication.php?id=")
                            || url.contains("/mep/add.php?e=")
                            || !currentDriver.findElements(
                                    By.linkText("Animal's Medication Prescriptions")).isEmpty() ? url : null;
                });
                if (responseUrl.contains("/mep/add.php?e=")) {
                    throw new IllegalStateException(
                            "The site rejected Prescription row " + (rowIndex + 1) + " for '" + catName
                                    + "' and returned " + responseUrl + ".");
                }
            } catch (TimeoutException e) {
                throw new IllegalStateException(
                        "Prescription row " + (rowIndex + 1) + " for '" + catName
                                + "' did not return to Medication Prescriptions. Current URL: "
                                + driver.getCurrentUrl(),
                        e);
            }
        }
    }

    static void addPrescriptionRecord(WebDriver driver, Sheet prescriptionSheet, int rowIndex,
            String catName) {
        int spreadsheetRow = rowIndex + 1;
        String medicationName = getCellValueAsString(getCell(prescriptionSheet, rowIndex, 1));
        String route = getCellValueAsString(getCell(prescriptionSheet, rowIndex, 2));
        String dosage = getCellValueAsString(getCell(prescriptionSheet, rowIndex, 3));
        String doseType = getCellValueAsString(getCell(prescriptionSheet, rowIndex, 4)).trim();
        LocalDate oneTimeDateNeeded = DateHelper.getCellValueAsDate(getCell(prescriptionSheet, rowIndex, 5));
        LocalDate oneTimeDateGiven = DateHelper.getCellValueAsDate(getCell(prescriptionSheet, rowIndex, 6));
        String administerer = getCellValueAsString(getCell(prescriptionSheet, rowIndex, 7));
        String createMAR = getCellValueAsString(getCell(prescriptionSheet, rowIndex, 8)).trim();
        String perDay = getCellValueAsString(getCell(prescriptionSheet, rowIndex, 9));
        String dateRange = getCellValueAsString(getCell(prescriptionSheet, rowIndex, 10));
        LocalDate recurringDateStart = DateHelper.getCellValueAsDate(getCell(prescriptionSheet, rowIndex, 11));
        LocalDate recurringDateEnd = DateHelper.getCellValueAsDate(getCell(prescriptionSheet, rowIndex, 12));
        String prescriber = getCellValueAsString(getCell(prescriptionSheet, rowIndex, 13));
        String vetClinic = getCellValueAsString(getCell(prescriptionSheet, rowIndex, 14));
        String genComments = getCellValueAsString(getCell(prescriptionSheet, rowIndex, 15));
        String hidComments = getCellValueAsString(getCell(prescriptionSheet, rowIndex, 16));

        if (!"One Time".equalsIgnoreCase(doseType) && !"Time Range".equalsIgnoreCase(doseType)) {
            throw new IllegalArgumentException(
                    "Prescriptions!E" + spreadsheetRow + " must be either 'One Time' or 'Time Range'.");
        }
        if (!createMAR.isEmpty() && !"Yes".equalsIgnoreCase(createMAR)
                && !"No".equalsIgnoreCase(createMAR)) {
            throw new IllegalArgumentException(
                    "Prescriptions!I" + spreadsheetRow + " must be either 'Yes' or 'No'.");
        }

        WebDriverWait formWait = new WebDriverWait(driver, Duration.ofSeconds(10));
        try {
            formWait.until(ExpectedConditions.visibilityOfElementLocated(By.name("m")));
        } catch (TimeoutException e) {
            throw new IllegalStateException(
                    "The medication dropdown did not appear for Prescriptions row " + spreadsheetRow
                            + ". Current URL: " + driver.getCurrentUrl(),
                    e);
        }

        selectPrescriptionMedication(driver, medicationName, spreadsheetRow);
        if (!route.isEmpty()) {
            FormHelper.selectByVisibleText(driver, "r", route);
        }
        FormHelper.enterText(driver, "d", dosage);

        if ("One Time".equalsIgnoreCase(doseType)) {
            driver.findElement(By.cssSelector("input[value='ot']")).click();
            if (!createMAR.isEmpty()) {
                FormHelper.setCheckbox(driver, "otcm", "Yes".equalsIgnoreCase(createMAR));
            }
            // +--------------------------------------------------------------------------+
            // | DEMO ONLY: the demo requires Date Needed; production does not.            |
            // | Keep this conditional field handling for demo testing, not as a general  |
            // | production requirement.                                                  |
            // +--------------------------------------------------------------------------+
            if (oneTimeDateNeeded != null) {
                DateHelper.enterDateText(driver, "otdn", oneTimeDateNeeded, false, true);
            }
            if (oneTimeDateGiven != null) {
                DateHelper.enterDateText(driver, "otdg", oneTimeDateGiven, false, true);
            }
            if (!administerer.isEmpty()) {
                FormHelper.selectByVisibleText(driver, "otp", administerer);
            }
        } else {
            driver.findElement(By.cssSelector("input[value='tb']")).click();
            if (!createMAR.isEmpty()) {
                FormHelper.setCheckbox(driver, "tbcm", "Yes".equalsIgnoreCase(createMAR));
            }
            selectNumericPrescriptionOption(
                    driver, "tbdf", perDay, 1, 10, "Prescriptions!J" + spreadsheetRow + " (How Often)");
            selectNumericPrescriptionOption(
                    driver, "tbgf", dateRange, 1, 30, "Prescriptions!K" + spreadsheetRow + " (Every)");

            if (recurringDateStart != null) {
                WebElement dateField = driver.findElement(By.name("tbsd"));
                DateHelper.selectDateFromPicker(driver, dateField, recurringDateStart);
            }
            if (recurringDateEnd != null) {
                WebElement dateField = driver.findElement(By.name("tbed"));
                DateHelper.selectDateFromPicker(driver, dateField, recurringDateEnd);
            }
        }

        if (!prescriber.isEmpty()) {
            FormHelper.selectByVisibleText(driver, "p", prescriber);
        }
        if (!vetClinic.isEmpty()) {
            FormHelper.selectByVisibleText(driver, "vc", vetClinic);
        }
        FormHelper.enterText(driver, "gc", genComments);
        FormHelper.enterText(driver, "hc", hidComments);
        driver.findElement(By.cssSelector("input#submit[type='submit'][value='Add']")).click();
    }

    static void selectPrescriptionMedication(WebDriver driver, String medicationName, int spreadsheetRow) {
        Select medicationDropdown = new Select(driver.findElement(By.name("m")));
        String normalizedMedicationName = medicationName.trim();
        WebElement matchingOption = medicationDropdown.getOptions().stream()
                .filter(option -> normalizedMedicationName.equalsIgnoreCase(option.getText().trim()))
                .findFirst()
                .orElse(null);

        if (matchingOption == null && "Amoxicillin".equalsIgnoreCase(normalizedMedicationName)) {
            matchingOption = medicationDropdown.getOptions().stream()
                    .filter(option -> "Amoxicilin".equalsIgnoreCase(option.getText().trim()))
                    .findFirst()
                    .orElse(null);
        }

        if (matchingOption == null) {
            String availableMedications = medicationDropdown.getOptions().stream()
                    .map(option -> option.getText().trim())
                    .filter(option -> !option.isEmpty())
                    .collect(Collectors.joining(", "));
            throw new IllegalArgumentException(
                    "Prescriptions!B" + spreadsheetRow + " medication '" + medicationName
                            + "' is not available. Available medications: [" + availableMedications + "].");
        }

        FormHelper.selectByValue(driver, "m", matchingOption.getAttribute("value"));
    }

    static void selectNumericPrescriptionOption(
            WebDriver driver, String fieldName, String value, int minimum, int maximum, String cellLabel) {
        String normalizedValue = value == null ? "" : value.trim();
        int numericValue;
        try {
            numericValue = Integer.parseInt(normalizedValue);
        } catch (NumberFormatException e) {
            throw new IllegalArgumentException(
                    cellLabel + " must be a whole number from " + minimum + " to " + maximum + ".", e);
        }
        if (numericValue < minimum || numericValue > maximum) {
            throw new IllegalArgumentException(
                    cellLabel + " must be a whole number from " + minimum + " to " + maximum + ".");
        }

        FormHelper.selectByValue(driver, fieldName, Integer.toString(numericValue));
    }
}

