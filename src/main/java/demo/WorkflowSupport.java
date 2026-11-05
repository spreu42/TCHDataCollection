package demo;

import helpers.DateHelper;
import helpers.AutocompleteHelper;
import helpers.FormHelper;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.StaleElementReferenceException;
import org.openqa.selenium.TimeoutException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import org.apache.poi.ss.usermodel.*;

import java.time.Duration;
import java.util.List;
import java.util.Map;


final class WorkflowSupport {
    public static void deleteRow(Sheet sheet, int rowIndex) {
        int lastRowNum = sheet.getLastRowNum();
        if (rowIndex >= 0 && rowIndex < lastRowNum) {
            sheet.shiftRows(rowIndex + 1, lastRowNum, -1);
        }
        if (rowIndex == lastRowNum) {
            Row removingRow = sheet.getRow(rowIndex);
            if (removingRow != null) {
                sheet.removeRow(removingRow);
            }
        }
    }

    static void selectPersonFromAutocomplete(
            WebDriver driver,
            String inputName,
            String suggestionListId,
            String linkedFieldName,
            String personName,
            String fieldLabel) {
        String normalizedName = personName == null ? "" : personName.trim();
        if (normalizedName.isEmpty()) {
            driver.findElement(By.name(inputName)).clear();
            return;
        }

        WebElement input = driver.findElement(By.name(inputName));
        input.clear();
        input.sendKeys(normalizedName.substring(0, 1));
        if (normalizedName.length() > 1) {
            input.sendKeys(normalizedName.substring(1));
        }

        By suggestions = By.cssSelector(
                "a[href^=\"javascript:setField('" + suggestionListId + "', '"
                        + inputName + "', '" + linkedFieldName + "',\"]");
        AutocompleteHelper.selectMatchingSuggestion(
                driver, suggestions, normalizedName, fieldLabel + " from the site's suggestions");
    }

    static String requireEnvironmentVariable(String name) {
        String value = System.getenv(name);
        if (value == null || value.trim().isEmpty()) {
            throw new IllegalStateException("Required environment variable is not set: " + name);
        }
        return value;
    }

    static boolean hasPrescriptionData(Sheet prescriptionSheet) {
        for (int rowIndex = 1; rowIndex <= prescriptionSheet.getLastRowNum(); rowIndex++) {
            String animalName = getCellValueAsString(getCell(prescriptionSheet, rowIndex, 0)).trim();
            if (animalName.isEmpty()) {
                continue;
            }
            if ("All Fields".equalsIgnoreCase(animalName)) {
                System.err.println("Skipping Prescriptions row " + (rowIndex + 1)
                        + ": the animal-name cell contains the 'All Fields' template placeholder.");
                continue;
            }
            return true;
        }
        return false;
    }

    static void openAnimalFromSearch(WebDriver driver, String animalName, String sourceCell) {
        String normalizedName = animalName == null ? "" : animalName.trim();
        if (normalizedName.isEmpty()) {
            throw new IllegalArgumentException("Animal name is blank in " + sourceCell + ".");
        }

        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
        By searchFieldLocator = By.id("s");
        wait.until(ExpectedConditions.visibilityOfElementLocated(searchFieldLocator));
        wait.until(currentDriver -> {
            try {
                WebElement searchField = currentDriver.findElement(searchFieldLocator);
                searchField.clear();
                searchField.sendKeys(normalizedName);
                return true;
            } catch (StaleElementReferenceException e) {
                return false;
            }
        });

        By suggestions = By.cssSelector("#sa a, #sa button, #sa [role='option'], #sa li");
        AutocompleteHelper.selectMatchingSuggestion(
                driver, suggestions, normalizedName, "Quick Search for " + sourceCell);
        wait.until(ExpectedConditions.presenceOfElementLocated(By.name("history")));
    }

    static Cell getCell(Sheet sheet, int rowIndex, int columnIndex) {
        Row row = sheet.getRow(rowIndex);
        return row == null ? null : row.getCell(columnIndex);
    }

    static boolean hasData(Sheet sheet) {
        Cell firstCell = getCell(sheet, 1, 0);
        return firstCell != null && !getCellValueAsString(firstCell).trim().isEmpty();
    }

    static String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateHelper.isDateCell(cell)) {
                    return DateHelper.formatCellDate(cell);
                } else {
                    double numValue = cell.getNumericCellValue();
                    if (numValue == (long) numValue) {
                        return String.format("%d", (long) numValue);
                    } else {
                        return String.format("%s", numValue);
                    }
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case BLANK:
                return "";
            default:
                return "";
        }
    }

    static boolean hasMedicalRecordData(Sheet sheet, String sheetName) {
        if (!hasData(sheet)) {
            return false;
        }

        String animalName = getCellValueAsString(getCell(sheet, 1, 0)).trim();
        if ("All Fields".equalsIgnoreCase(animalName)) {
            System.err.println("Skipping " + sheetName
                    + ": A2 contains the 'All Fields' template placeholder, not an animal name.");
            return false;
        }

        return true;
    }

    static WebElement openAnimalHistoryAddForm(
            WebDriver driver,
            String animalName,
            String historyName,
            String historyRoute,
            String addRecordTitle) {
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
        Select historyDropdown = new Select(
                wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector("select.history"))));
        WebElement historyOption = historyDropdown.getOptions().stream()
                .filter(option -> historyName.equals(option.getText().trim()))
                .findFirst()
                .orElseThrow(() -> new IllegalStateException(
                        "The history dropdown has no " + historyName
                                + " option for animal '" + animalName + "'."));
        String historyUrl = historyOption.getAttribute("value");
        if (historyUrl == null || !historyUrl.startsWith(historyRoute + "?id=")) {
            throw new IllegalStateException(
                    "The " + historyName + " option has an unexpected destination for animal '"
                            + animalName + "': " + historyUrl);
        }
        FormHelper.selectByVisibleText(driver, "history", historyName);

        By addRecordControl = By.cssSelector(
                "[title='" + addRecordTitle + "'], [alt='" + addRecordTitle + "']");
        try {
            wait.until(ExpectedConditions.urlContains(historyUrl));
            return wait.until(ExpectedConditions.elementToBeClickable(addRecordControl));
        } catch (TimeoutException e) {
            printPageSnapshot(driver, historyName + " navigation for '" + animalName + "'");
            throw new IllegalStateException(
                    "Opened " + historyName + " for animal '" + animalName
                            + "', but the '" + addRecordTitle + "' control did not appear. Current URL: "
                            + driver.getCurrentUrl(),
                    e);
        }
    }

    static boolean hasCheckInCatNames(Sheet checkinSheet) {
        Cell namesCell = getCell(checkinSheet, 1, 1);
        return namesCell != null && !getCellValueAsString(namesCell).trim().isEmpty();
    }

    static void logCheckInField(WebDriver driver, By locator, String fieldName, String expectedValue) {
        List<WebElement> fields = driver.findElements(locator);
        if (fields.isEmpty()) {
            System.out.printf("[Check-In debug] field=%s, expected=%s, actual=NOT FOUND%n",
                    fieldName, expectedValue);
            return;
        }

        WebElement field = fields.get(0);
        Object actualValue = ((JavascriptExecutor) driver).executeScript(
                "var el = arguments[0];" +
                        "return JSON.stringify({" +
                        "tag: el.tagName," +
                        "type: el.type || null," +
                        "value: el.value === undefined ? null : el.value," +
                        "checked: ('checked' in el) ? el.checked : null," +
                        "selectedOptions: el.tagName === 'SELECT'" +
                        "  ? Array.from(el.selectedOptions).map(function(option) { return option.text; })" +
                        "  : null" +
                        "});",
                field);
        System.out.printf("[Check-In debug] field=%s, expected=%s, actual=%s%n",
                fieldName, expectedValue, actualValue);
    }

    static void printPageSnapshot(WebDriver driver, String label) {
        Object text = ((JavascriptExecutor) driver).executeScript(
                "var c = document.getElementById('content') || document.body;" +
                        "return c.innerText.substring(0, 5000);");
        System.err.println("[Page debug] " + label + ": title='" + driver.getTitle()
                + "', url=" + driver.getCurrentUrl());
        System.err.println("[Page debug] content text:\n" + text);
    }
}

