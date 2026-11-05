package tests;

import org.openqa.selenium.By;
import org.openqa.selenium.Alert;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.NoAlertPresentException;
import org.openqa.selenium.TimeoutException;
import org.openqa.selenium.UnhandledAlertException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.time.Duration;
import java.time.LocalDate;
import java.time.YearMonth;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.Locale;
import java.text.SimpleDateFormat;
import java.util.stream.Collectors;

public class FurstTest {
    public static void main(String[] args) throws IOException {
        String username = requireEnvironmentVariable("TCH_USER");
        String password = requireEnvironmentVariable("TCH_PASSWORD");
        File workbookFile = new File("src/main/resources/DCWorkbook.xlsm");
        try (FileInputStream input = new FileInputStream(workbookFile);
             Workbook workbook = new XSSFWorkbook(input)) {
            Sheet newCatSheet = workbook.getSheet("New Intake - Add");
            Sheet checkinSheet = workbook.getSheet("New Intake - Check-In");
            Sheet vaccSheet = workbook.getSheet("Vaccinations");
            Sheet medCheckSheet = workbook.getSheet("Medical Check-Up");
            Sheet procedureSheet = workbook.getSheet("Medical Procedure");
            Sheet prescriptionSheet = workbook.getSheet("Prescriptions");

            WebDriver driver = new ChromeDriver();
            try {
                // Production site; keep disabled while testing.
                // driver.get("https://ishelter.ishelters.com/");
                driver.get("https://ishelter.demoshelters.com/login.php");

                WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
                wait.until(ExpectedConditions.presenceOfElementLocated(By.id("email")));
                driver.findElement(By.id("email")).sendKeys(username);
                driver.findElement(By.id("password")).sendKeys(password);
                driver.findElement(By.id("submit")).click();
                wait.until(ExpectedConditions.elementToBeClickable(
                        By.cssSelector("a[href='/as/']"))).click();

                String newlyAddedAnimalName = null;
                if (hasData(newCatSheet)) {
                    newlyAddedAnimalName = getCellValueAsString(newCatSheet.getRow(1).getCell(0)).trim();
                    addNewCat(driver, newCatSheet);
                }
                boolean checkInCompleted = false;
                if (hasCheckInCatNames(checkinSheet)) {
                    checkInCompleted = createCheckInRecord(driver, checkinSheet, newlyAddedAnimalName);
                } else if (hasData(checkinSheet)) {
                    System.err.println(
                            "Skipping check-in: the cat count is filled, but the cat-name list is blank.");
                }
                if (checkInCompleted && isCheckInOptionYes(checkinSheet, 16) && hasData(vaccSheet)) {
                    addVaccinations(driver, vaccSheet);
                }
                if (checkInCompleted && isCheckInOptionYes(checkinSheet, 17) && hasData(medCheckSheet)) {
                    addMedCheck(driver, medCheckSheet);
                }
                if (checkInCompleted && isCheckInOptionYes(checkinSheet, 17) && hasData(procedureSheet)) {
                    addProcedure(driver, procedureSheet);
                }
                if (checkInCompleted && isCheckInOptionYes(checkinSheet, 18) && hasData(prescriptionSheet)) {
                    addPrescriptions(driver, prescriptionSheet);
                }
            } finally {
                driver.quit();
            }
        }
    }

    private static boolean hasData(Sheet sheet) {
        Cell firstCell = getCell(sheet, 1, 0);
        return firstCell != null && !getCellValueAsString(firstCell).trim().isEmpty();
    }

    private static boolean hasCheckInCatNames(Sheet checkinSheet) {
        Cell namesCell = getCell(checkinSheet, 1, 1);
        return namesCell != null && !getCellValueAsString(namesCell).trim().isEmpty();
    }

    private static boolean isCheckInOptionYes(Sheet checkinSheet, int columnIndex) {
        Cell optionCell = getCell(checkinSheet, 1, columnIndex);
        return optionCell != null && "Yes".equalsIgnoreCase(getCellValueAsString(optionCell).trim());
    }

    private static Cell getCell(Sheet sheet, int rowIndex, int columnIndex) {
        Row row = sheet.getRow(rowIndex);
        return row == null ? null : row.getCell(columnIndex);
    }

    private static String requireEnvironmentVariable(String name) {
        String value = System.getenv(name);
        if (value == null || value.trim().isEmpty()) {
            throw new IllegalStateException("Required environment variable is not set: " + name);
        }
        return value;
    }

    private static void selectDateFromPicker(WebDriver driver, WebElement dateField, LocalDate date) {
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
        By calendar = By.cssSelector("table.ui-datepicker-calendar");
        By title = By.cssSelector(".ui-datepicker-title");
        By monthSelect = By.cssSelector(".ui-datepicker-title select.ui-datepicker-month");
        By yearSelect = By.cssSelector(".ui-datepicker-title select.ui-datepicker-year");
        String targetInputId = dateField.getDomAttribute("id");
        if (targetInputId == null || targetInputId.isEmpty()) {
            throw new IllegalStateException("Date of birth input is missing its id.");
        }
        boolean calendarVisible = !driver.findElements(calendar).isEmpty()
                && driver.findElement(calendar).isDisplayed();
        if (calendarVisible && !targetInputId.equals(getDatePickerInputId(driver))) {
            hideDatePicker(driver);
            wait.until(ExpectedConditions.invisibilityOfElementLocated(calendar));
            calendarVisible = false;
        }
        if (!calendarVisible) {
            dateField.click();
        }
        wait.until(ExpectedConditions.visibilityOfElementLocated(calendar));
        wait.until(currentDriver -> targetInputId.equals(getDatePickerInputId(currentDriver)));

        if (!driver.findElements(monthSelect).isEmpty() && !driver.findElements(yearSelect).isEmpty()) {
            new Select(wait.until(ExpectedConditions.visibilityOfElementLocated(yearSelect)))
                    .selectByValue(Integer.toString(date.getYear()));
            new Select(wait.until(ExpectedConditions.visibilityOfElementLocated(monthSelect)))
                    .selectByValue(Integer.toString(date.getMonthValue() - 1));
            wait.until(currentDriver ->
                    Integer.toString(date.getYear()).equals(
                            currentDriver.findElement(yearSelect).getDomProperty("value"))
                            && Integer.toString(date.getMonthValue() - 1).equals(
                            currentDriver.findElement(monthSelect).getDomProperty("value")));
        } else {
            DateTimeFormatter monthYearFormat = DateTimeFormatter.ofPattern("MMMM uuuu", Locale.ENGLISH);
            String displayedMonthYear = driver.findElement(title).getText().trim();
            if (!displayedMonthYear.matches("[A-Za-z]+\\s+\\d{4}")) {
                throw new IllegalStateException(
                        "Date picker did not expose month/year selectors or a readable month/year title: "
                                + displayedMonthYear);
            }
            YearMonth targetMonth = YearMonth.from(date);
            YearMonth displayedMonth = YearMonth.parse(displayedMonthYear, monthYearFormat);
            for (int attempt = 0; attempt < 240 && !displayedMonth.equals(targetMonth); attempt++) {
                By navigation = By.cssSelector(displayedMonth.isAfter(targetMonth)
                        ? "a.ui-datepicker-prev"
                        : "a.ui-datepicker-next");
                wait.until(ExpectedConditions.elementToBeClickable(navigation)).click();
                String previousMonthYear = displayedMonthYear;
                wait.until(currentDriver ->
                        !currentDriver.findElement(title).getText().trim().equals(previousMonthYear));
                displayedMonthYear = driver.findElement(title).getText().trim();
                displayedMonth = YearMonth.parse(displayedMonthYear, monthYearFormat);
            }
            if (!displayedMonth.equals(targetMonth)) {
                throw new IllegalStateException("Could not navigate the date picker to " + targetMonth + ".");
            }
        }

        By dayLink = By.xpath("//table[contains(@class,'ui-datepicker-calendar')]"
                + "//td[not(contains(@class,'ui-datepicker-other-month'))]"
                + "/a[normalize-space(text())='" + date.getDayOfMonth() + "']");
        wait.until(ExpectedConditions.elementToBeClickable(dayLink)).click();

        String enteredDate = dateField.getDomProperty("value");
        if (!date.toString().equals(enteredDate)) {
            printDatePickerDetails(driver, dateField, date);
            throw new IllegalStateException(
                    "Could not select date of birth. Expected " + date
                            + " but the field contains '" + enteredDate + "'.");
        }
        System.out.printf("[DOB debug] method=date-picker, accepted=true, value=%s%n", enteredDate);
    }

    private static String getDatePickerInputId(WebDriver driver) {
        Object inputId = ((JavascriptExecutor) driver).executeScript(
                "var cell = document.querySelector('.ui-datepicker-calendar td[onclick]');" +
                        "var handler = cell ? cell.getAttribute('onclick') : null;" +
                        "var match = handler ? handler.match(/_selectDay\\('#([^']+)'/) : null;" +
                        "return match ? match[1] : null;");
        return inputId == null ? null : inputId.toString();
    }

    private static void printDatePickerDetails(WebDriver driver, WebElement dateField, LocalDate expectedDate) {
        Object details = ((JavascriptExecutor) driver).executeScript(
                "var field = arguments[0];" +
                        "var calendars = Array.from(document.querySelectorAll('.ui-datepicker'));" +
                        "return JSON.stringify({" +
                        "url: window.location.href," +
                        "field: field ? {" +
                        "  outerHTML: field.outerHTML," +
                        "  value: field.value," +
                        "  parentHTML: field.parentElement ? field.parentElement.outerHTML : null" +
                        "} : null," +
                        "calendarCount: calendars.length," +
                        "calendars: calendars.map(function(calendar) {" +
                        "  var title = calendar.querySelector('.ui-datepicker-title');" +
                        "  var table = calendar.querySelector('table.ui-datepicker-calendar');" +
                        "  var monthSelect = calendar.querySelector('.ui-datepicker-month');" +
                        "  var yearSelect = calendar.querySelector('.ui-datepicker-year');" +
                        "  return {" +
                        "    id: calendar.id," +
                        "    display: getComputedStyle(calendar).display," +
                        "    visibility: getComputedStyle(calendar).visibility," +
                        "    className: calendar.className," +
                        "    titleText: title ? title.innerText : null," +
                        "    titleHTML: title ? title.outerHTML : null," +
                        "    monthSelectHTML: monthSelect ? monthSelect.outerHTML : null," +
                        "    yearSelectHTML: yearSelect ? yearSelect.outerHTML : null," +
                        "    calendarTableHTML: table ? table.outerHTML : null" +
                        "  };" +
                        "})" +
                        "})",
                dateField);
        System.err.println("[DOB debug] date-picker DOM details: " + details);
        System.err.println("[DOB debug] date-picker input id: " + getDatePickerInputId(driver));
        System.err.println("[DOB debug] expected date: " + expectedDate);
    }

    private static void hideDatePicker(WebDriver driver) {
        Object hidden = ((JavascriptExecutor) driver).executeScript(
                "var picker = window.DP_jQuery || window.jQuery;" +
                        "if (!picker || !picker.datepicker) return false;" +
                        "picker.datepicker._hideDatepicker();" +
                        "return true;");
        if (!Boolean.TRUE.equals(hidden)) {
            throw new IllegalStateException("Could not close the open jQuery UI date picker.");
        }
    }

    private static void printPageSnapshot(WebDriver driver, String label) {
        Object text = ((JavascriptExecutor) driver).executeScript(
                "var c = document.getElementById('content') || document.body;" +
                        "return c.innerText.substring(0, 1500);");
        System.err.println("[Page debug] " + label + ": title='" + driver.getTitle()
                + "', url=" + driver.getCurrentUrl());
        System.err.println("[Page debug] content text:\n" + text);
    }

    private static void logCheckInField(WebDriver driver, By locator, String fieldName, String expectedValue) {
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

    private static void selectPersonFromAutocomplete(
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
        try {
            WebElement matchingSuggestion = new WebDriverWait(driver, Duration.ofSeconds(10))
                    .until(currentDriver -> currentDriver.findElements(suggestions).stream()
                            .filter(WebElement::isDisplayed)
                            .filter(option -> suggestionMatchesName(option.getText(), normalizedName))
                            .findFirst()
                            .orElse(null));
            matchingSuggestion.click();
        } catch (TimeoutException e) {
            List<String> availableSuggestions = driver.findElements(suggestions).stream()
                    .filter(WebElement::isDisplayed)
                    .map(option -> option.getText().trim())
                    .collect(Collectors.toList());
            throw new IllegalStateException(
                    "Could not select " + fieldLabel + " '" + normalizedName
                            + "' from the site's suggestions. Available suggestions: "
                            + availableSuggestions,
                    e);
        }
    }

    private static boolean suggestionMatchesName(String suggestionText, String expectedName) {
        String normalizedSuggestion = suggestionText == null ? "" : suggestionText.trim();
        return normalizedSuggestion.equalsIgnoreCase(expectedName)
                || normalizedSuggestion.regionMatches(
                        true, 0, expectedName + " (", 0, expectedName.length() + 2);
    }

    private static void openAnimalFromSearch(WebDriver driver, String animalName, String sourceCell) {
        String normalizedName = animalName == null ? "" : animalName.trim();
        if (normalizedName.isEmpty()) {
            throw new IllegalArgumentException("Animal name is blank in " + sourceCell + ".");
        }

        WebElement searchField = driver.findElement(By.id("s"));
        searchField.clear();
        searchField.sendKeys(normalizedName);
        submitSearchForm(searchField);

        try {
            WebElement result = new WebDriverWait(driver, Duration.ofSeconds(10)).until(currentDriver ->
                    currentDriver.findElements(By.cssSelector("a.animal")).stream()
                            .filter(link -> normalizedName.equals(link.getText().trim()))
                            .findFirst()
                            .orElse(null));
            result.click();
        } catch (TimeoutException e) {
            printPageSnapshot(driver, "animal search for '" + normalizedName + "'");
            List<WebElement> links = driver.findElements(By.cssSelector("a.animal"));
            System.err.println("[Search debug] a.animal links found: " + links.size());
            for (WebElement link : links) {
                System.err.println("[Search debug]   '" + link.getText().trim() + "'");
            }
            throw new IllegalStateException(
                    "No animal named '" + normalizedName + "' was found for " + sourceCell
                            + ". Check that the spreadsheet cell contains an animal name.",
                    e);
        }
    }

    private static void submitSearchForm(WebElement searchField) {
        WebElement searchForm = searchField.findElement(By.xpath("./ancestor::form[1]"));
        WebElement submitButton = searchForm
                .findElements(By.cssSelector("button[type='submit'], input[type='submit']")).stream()
                .filter(WebElement::isDisplayed)
                .filter(WebElement::isEnabled)
                .findFirst()
                .orElseThrow(() -> new IllegalStateException(
                        "Could not find an enabled submit button in the animal search form."));
        submitButton.click();
    }

    private static void openCheckInFormForExistingAnimal(WebDriver driver, String animalName) {
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
        Select historyDropdown = new Select(
                wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector("select.history"))));
        WebElement checkInHistoryOption = historyDropdown.getOptions().stream()
                .filter(option -> "Check-In History".equals(option.getText().trim()))
                .findFirst()
                .orElseThrow(() -> new IllegalStateException(
                        "The history dropdown has no Check-In History option for animal '"
                                + animalName + "'."));
        String historyUrl = checkInHistoryOption.getAttribute("value");
        if (historyUrl == null || !historyUrl.startsWith("check.php?id=")) {
            throw new IllegalStateException(
                    "The Check-In History option has an unexpected destination for animal '"
                            + animalName + "': " + historyUrl);
        }

        historyDropdown.selectByVisibleText("Check-In History");
        try {
            wait.until(ExpectedConditions.urlContains(historyUrl));
        } catch (TimeoutException e) {
            printPageSnapshot(driver, "Check-In History navigation for '" + animalName + "'");
            throw new IllegalStateException(
                    "Selecting Check-In History did not open its destination for animal '"
                            + animalName + "'. Expected URL to contain " + historyUrl,
                    e);
        }

        By addCheckInLink = By.xpath(
                "//a[.//img[@title='Add a Check-In' or @alt='Add a Check-In']]");
        try {
            wait.until(ExpectedConditions.elementToBeClickable(addCheckInLink)).click();
        } catch (TimeoutException e) {
            printPageSnapshot(driver, "Add Check-In navigation for '" + animalName + "'");
            throw new IllegalStateException(
                    "Opened Check-In History for animal '" + animalName
                            + "', but the Add a Check-In link was not found.",
                    e);
        }
    }

    private static String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return new SimpleDateFormat("MM/dd/yyyy").format(cell.getDateCellValue());
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

    private static void addNewCat(WebDriver driver, Sheet newCatSheet) {
        // Implementation for adding a new cat based on the data in newCatSheet
        
        //Gather all filled fields for a new cat
        String nameofCat = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(0) != null) ? getCellValueAsString(newCatSheet.getRow(1).getCell(0)) : "";
        String AliasofCat = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(1) != null) ? getCellValueAsString(newCatSheet.getRow(1).getCell(1)) : "";
        String species = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(2) != null) ? getCellValueAsString(newCatSheet.getRow(1).getCell(2)) : "";
        String breed = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(3) != null) ? getCellValueAsString(newCatSheet.getRow(1).getCell(3)) : "";
        String catCode = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(4) != null) ? getCellValueAsString(newCatSheet.getRow(1).getCell(4)) : "";
        String catType = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(5) != null) ? getCellValueAsString(newCatSheet.getRow(1).getCell(5)) : "";
        String readyToAdopt = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(6) != null) ? getCellValueAsString(newCatSheet.getRow(1).getCell(6)) : "";
        String webDisplay = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(7) != null) ? getCellValueAsString(newCatSheet.getRow(1).getCell(7)) : "";
        String catStatus = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(8) != null) ? getCellValueAsString(newCatSheet.getRow(1).getCell(8)) : "";
        LocalDate statusChangeDate = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(9) != null) ? getCellValueAsDate(newCatSheet.getRow(1).getCell(9)) : null;
        String gender = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(10) != null) ? getCellValueAsString(newCatSheet.getRow(1).getCell(10)) : "";
        String catColor = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(11) != null) ? getCellValueAsString(newCatSheet.getRow(1).getCell(11)) : "";
        String chipType = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(12) != null) ? getCellValueAsString(newCatSheet.getRow(1).getCell(12)) : "";
        String chipNumber = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(13) != null) ? getCellValueAsString(newCatSheet.getRow(1).getCell(13)) : "";
        String tagNumber = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(14) != null) ? getCellValueAsString(newCatSheet.getRow(1).getCell(14)) : "";
        String DOBest = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(15) != null) ? getCellValueAsString(newCatSheet.getRow(1).getCell(15)) : "";
        LocalDate DOB = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(16) != null) ? getCellValueAsDate(newCatSheet.getRow(1).getCell(16)) :null;
        LocalDate altered = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(17) != null) ? getCellValueAsDate(newCatSheet.getRow(1).getCell(17)) : null;
        String genComments = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(18) != null) ? getCellValueAsString(newCatSheet.getRow(1).getCell(18)) : "";
        String hidCommments = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(19) != null) ? getCellValueAsString(newCatSheet.getRow(1).getCell(19)) : "";
        String distinctFeatures = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(20) != null) ? getCellValueAsString(newCatSheet.getRow(1).getCell(20)) : "";
        String shortBio = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(21) != null) ? getCellValueAsString(newCatSheet.getRow(1).getCell(21)) : "";
        String longBio = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(22) != null) ? getCellValueAsString(newCatSheet.getRow(1).getCell(22)) : "";
        String fearfulStat = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(23) != null) ? getCellValueAsString(newCatSheet.getRow(1).getCell(23)) : "";
        String childrenStat = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(24) != null) ? getCellValueAsString(newCatSheet.getRow(1).getCell(24)) : "";
        String animalStat = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(25) != null) ? getCellValueAsString(newCatSheet.getRow(1).getCell(25)) : "";
        String socialStat = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(26) != null) ? getCellValueAsString(newCatSheet.getRow(1).getCell(26)) : "";
        String declawedStat = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(27) != null) ? getCellValueAsString(newCatSheet.getRow(1).getCell(27)) : "";
        String litterTrainStat = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(28) != null) ? getCellValueAsString(newCatSheet.getRow(1).getCell(28)) : "";
        String specialNeedsStat = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(29) != null) ? getCellValueAsString(newCatSheet.getRow(1).getCell(29)) : "";

        if (hasData(newCatSheet)) {
            By addAnimalLink = By.linkText("Add Animal");
            try {
                new WebDriverWait(driver, Duration.ofSeconds(15))
                        .until(ExpectedConditions.elementToBeClickable(addAnimalLink))
                        .click();
            } catch (TimeoutException e) {
                throw new IllegalStateException(
                        "Could not find the 'Add Animal' link after login. Current page: "
                                + driver.getTitle() + " (" + driver.getCurrentUrl() + ")",
                        e);
            }
            driver.findElement(By.name("n")).sendKeys(nameofCat);
            driver.findElement(By.name("as")).sendKeys(AliasofCat);
            driver.findElement(By.cssSelector("div[id='content'] select[name='s']")).sendKeys(species);
            driver.findElement(By.name("b")).sendKeys(breed);
            driver.findElement(By.name("c")).sendKeys(catCode);
            driver.findElement(By.name("at")).sendKeys(catType);
            if ("Yes".equals(readyToAdopt)) {
                driver.findElement(By.name("rta")).click();
            }
            if ("Yes".equals(webDisplay)) {
                driver.findElement(By.name("sw")).click();
            }
            driver.findElement(By.name("st")).sendKeys(catStatus);
            if (statusChangeDate != null) {
                driver.findElement(By.name("std")).click();
                driver.findElement(By.name("std")).sendKeys(Keys.chord(Keys.CONTROL + "a"), Keys.DELETE);
                driver.findElement(By.name("std")).sendKeys(statusChangeDate.toString());
            }
            driver.findElement(By.name("g")).sendKeys(gender);
            driver.findElement(By.name("pc")).sendKeys(catColor);
            driver.findElement(By.name("mt")).sendKeys(chipType);
            driver.findElement(By.name("mn")).sendKeys(chipNumber);
            driver.findElement(By.name("tn")).sendKeys(tagNumber);
            
            if (DOB != null) {
                WebElement dobField = driver.findElement(By.name("bd"));
                selectDateFromPicker(driver, dobField, DOB);
            }

            if (altered != null) {
                driver.findElement(By.name("nd")).click();
                driver.findElement(By.name("nd")).sendKeys(Keys.chord(Keys.CONTROL + "a"), Keys.DELETE);
                driver.findElement(By.name("nd")).sendKeys(altered.toString());
                driver.findElement(By.name("nd")).sendKeys(Keys.TAB);
            }

            driver.findElement(By.name("gc")).sendKeys(genComments);
            driver.findElement(By.name("hc")).sendKeys(hidCommments);
            driver.findElement(By.name("df")).sendKeys(distinctFeatures);
            driver.findElement(By.name("sbio")).sendKeys(shortBio);
            driver.findElement(By.name("lbio")).sendKeys(longBio);
            
            //driver.findElement(By.cssSelector("label[for='e']")).click();

            if ("Yes".equals(DOBest)) {
                ((JavascriptExecutor) driver).executeScript("document.getElementsByName('e')[0].click();");
            }

            if ("Yes".equals(fearfulStat)) {
                driver.findElement(By.cssSelector("[value='8309']")).click();
            }
            if ("Yes".equals(childrenStat)) {
                driver.findElement(By.cssSelector("[value='8307']")).click();
            }
            if ("Yes".equals(animalStat)) {
                driver.findElement(By.cssSelector("[value='8310']")).click();
            }
            if ("Yes".equals(socialStat)) {
                driver.findElement(By.cssSelector("[value='8308']")).click();
            }
            if ("Yes".equals(declawedStat)) {
                driver.findElement(By.name("aa6154")).click();
            }
            if ("Yes".equals(litterTrainStat)) {
                driver.findElement(By.name("aa6155")).click();
            }
            if ("Yes".equals(specialNeedsStat)) {
                driver.findElement(By.name("aa6156")).click();
            }
            //WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
            //wait.until(ExpectedConditions.elementToBeClickable(By.id("justkidding")));
            
            driver.findElement(By.cssSelector("div[id='content'] input[id='submit']")).click();
            /*!!! Input a check to see if it takes you to the "Similar Animals" page !!!*/
            /*!!! If it does, need to count how many have the same name(*) and then +1 whatever the number is append to the end !!!*/
            // Row clearing is disabled so the same test data can be reused.
            // deleteRow(newCatSheet, 1);
        }
    }

    private static boolean createCheckInRecord(
            WebDriver driver, Sheet checkinSheet, String newlyAddedAnimalName) {
        // Implementation for creating a check-in record based on the data in checkinSheet
        String numberCats = getCellValueAsString(getCell(checkinSheet, 1, 0));
        int numCats;
        try {
            numCats = Integer.parseInt(numberCats.trim());
        } catch (NumberFormatException e) {
            throw new IllegalArgumentException(
                    "Check-in sheet has cat names but the cat count is blank or not a whole number.", e);
        }
        String catNames = getCellValueAsString(getCell(checkinSheet, 1, 1));
        List<String> catList = Arrays.stream(catNames.split(";"))
                .map(String::trim)
                .filter(name -> !name.isEmpty())
                .collect(Collectors.toList());
        if (numCats <= 0 || catList.size() != numCats) {
            throw new IllegalArgumentException(
                    "Check-in sheet lists " + numCats + " cats but contains " + catList.size()
                            + " non-empty names.");
        }
        String broughtBy = getCellValueAsString(getCell(checkinSheet, 1, 2));
        String prevOwner = getCellValueAsString(getCell(checkinSheet, 1, 3));
        String prevAgency = getCellValueAsString(getCell(checkinSheet, 1, 4));
        String surrReason = getCellValueAsString(getCell(checkinSheet, 1, 5));
        String checkinType = getCellValueAsString(getCell(checkinSheet, 1, 6));
        Cell checkinDateCell = getCell(checkinSheet, 1, 7);
        LocalDate checkinDate = checkinDateCell == null ? null : getCellValueAsDate(checkinDateCell);
        String checker = getCellValueAsString(getCell(checkinSheet, 1, 8));
        String sheltAnCounts = getCellValueAsString(getCell(checkinSheet, 1, 9));
        String genComments = getCellValueAsString(getCell(checkinSheet, 1, 10));
        String hidComments = getCellValueAsString(getCell(checkinSheet, 1, 11));
        String move = getCellValueAsString(getCell(checkinSheet, 1, 12));
        String currentlyWith = getCellValueAsString(getCell(checkinSheet, 1, 13));
        String location = getCellValueAsString(getCell(checkinSheet, 1, 14));
        String subLocation = getCellValueAsString(getCell(checkinSheet, 1, 15));
        LocalDate today = LocalDate.now();

        if (hasCheckInCatNames(checkinSheet)) {
            if (newlyAddedAnimalName != null && newlyAddedAnimalName.equals(catList.get(0))) {
                By directCheckInLink = By.xpath(
                        "//a[normalize-space(.)='Check-In' and starts-with(@href, '/chp/add.php?id=')]");
                try {
                    new WebDriverWait(driver, Duration.ofSeconds(15))
                            .until(ExpectedConditions.elementToBeClickable(directCheckInLink))
                            .click();
                } catch (TimeoutException e) {
                    throw new IllegalStateException(
                            "The new animal was added, but its direct Check-In link was not found. Current page: "
                                    + driver.getTitle() + " (" + driver.getCurrentUrl() + ")",
                            e);
                }
            } else {
                openAnimalFromSearch(driver, catList.get(0), "New Intake - Check-In!B2");
                openCheckInFormForExistingAnimal(driver, catList.get(0));
            }

            WebElement firstAnimalField = new WebDriverWait(driver, Duration.ofSeconds(10))
                    .until(ExpectedConditions.visibilityOfElementLocated(By.name("aa1")));
            String selectedFirstAnimal = firstAnimalField.getDomProperty("value");
            if (!catList.get(0).equals(selectedFirstAnimal == null ? "" : selectedFirstAnimal.trim())) {
                throw new IllegalStateException(
                        "The Check-In form's first animal is '" + selectedFirstAnimal
                                + "', but the first name in New Intake - Check-In!B2 is '"
                                + catList.get(0) + "'.");
            }
            logCheckInField(driver, By.name("aa1"), "first animal", catList.get(0));

            // The first animal is already selected; add and select only the remaining names.
            for (int i = 1; i <= numCats - 1; i++) {
                driver.findElement(By.cssSelector("a[title='Add an additional animal to be checked in']")).click();
                int fieldNumber = i + 1;
                By additionalAnimalField = By.name("aa" + fieldNumber);
                new WebDriverWait(driver, Duration.ofSeconds(10))
                        .until(ExpectedConditions.visibilityOfElementLocated(additionalAnimalField));
                selectPersonFromAutocomplete(
                        driver,
                        "aa" + fieldNumber,
                        "aaa" + fieldNumber,
                        "a" + fieldNumber,
                        catList.get(i),
                        "Check-In animal " + fieldNumber);
                logCheckInField(driver, additionalAnimalField,
                        "additional animal " + fieldNumber, catList.get(i));
            }

            selectPersonFromAutocomplete(driver, "bb", "bba", "b", broughtBy, "Brought in By");
            logCheckInField(driver, By.name("bb"), "bb (brought by)", broughtBy);
            selectPersonFromAutocomplete(driver, "po", "poa", "p", prevOwner, "Previous Owner");
            logCheckInField(driver, By.name("po"), "po (previous owner)", prevOwner);
            driver.findElement(By.name("j")).sendKeys(prevAgency);
            logCheckInField(driver, By.name("j"), "j (previous agency)", prevAgency);
            driver.findElement(By.name("sr")).sendKeys(surrReason);
            logCheckInField(driver, By.name("sr"), "sr (surrender reason)", surrReason);
            driver.findElement(By.name("ct")).sendKeys(checkinType);
            logCheckInField(driver, By.name("ct"), "ct (check-in type)", checkinType);
            if (checkinDate != null) {
                driver.findElement(By.name("cd")).click();
                driver.findElement(By.name("cd")).sendKeys(Keys.chord(Keys.CONTROL + "a"), Keys.DELETE);
                driver.findElement(By.name("cd")).sendKeys(checkinDate.toString());
                driver.findElement(By.name("cd")).sendKeys(Keys.TAB);
                logCheckInField(driver, By.name("cd"), "cd (check-in date)", checkinDate.toString());
            } else {
                driver.findElement(By.name("cd")).click();
                driver.findElement(By.name("cd")).sendKeys(Keys.chord(Keys.CONTROL + "a"), Keys.DELETE);
                driver.findElement(By.name("cd")).sendKeys(today.toString());
                driver.findElement(By.name("cd")).sendKeys(Keys.TAB);
                logCheckInField(driver, By.name("cd"), "cd (check-in date)", today.toString());
            }
            driver.findElement(By.name("cb")).sendKeys(checker);
            logCheckInField(driver, By.name("cb"), "cb (checker)", checker);
            driver.findElement(By.name("sacs")).sendKeys(sheltAnCounts);
            logCheckInField(driver, By.name("sacs"), "sacs (shelter animal counts)", sheltAnCounts);
            driver.findElement(By.name("gc")).sendKeys(genComments);
            logCheckInField(driver, By.name("gc"), "gc (general comments)", genComments);
            driver.findElement(By.name("hc")).sendKeys(hidComments);
            logCheckInField(driver, By.name("hc"), "hc (hidden comments)", hidComments);
            
            //Process a move
            By moveCheckbox = By.name("mrc");
            if ("Yes".equals(move)) {
                driver.findElement(moveCheckbox).click();
                logCheckInField(driver, moveCheckbox, "mrc (move requested)", move);
                selectPersonFromAutocomplete(
                        driver, "mrcw", "mrcwa", "mrp", currentlyWith, "Currently With");
                logCheckInField(driver, By.name("mrcw"), "mrcw (currently with)", currentlyWith);
                driver.findElement(By.name("mrl")).sendKeys(location);
                logCheckInField(driver, By.name("mrl"), "mrl (location)", location);
                driver.findElement(By.name("mrsl")).sendKeys(subLocation);
                logCheckInField(driver, By.name("mrsl"), "mrsl (sub-location)", subLocation);
            } else {
                logCheckInField(driver, moveCheckbox, "mrc (move requested)", move);
            }
            
            By addButton = By.cssSelector("input[type='submit'][name='submit'][value='Add']");
            String formUrl = driver.getCurrentUrl();
            driver.findElement(addButton).click();
            try {
                String submissionResult = new WebDriverWait(driver, Duration.ofSeconds(15))
                        .until(currentDriver -> {
                            try {
                                Alert alert = currentDriver.switchTo().alert();
                                return "alert:" + alert.getText();
                            } catch (NoAlertPresentException ignored) {
                                try {
                                    if (!formUrl.equals(currentDriver.getCurrentUrl())
                                            || currentDriver.findElements(addButton).stream()
                                                    .noneMatch(WebElement::isDisplayed)) {
                                        return "submitted";
                                    }
                                } catch (UnhandledAlertException alertOpened) {
                                    try {
                                        return "alert:" + currentDriver.switchTo().alert().getText();
                                    } catch (NoAlertPresentException alertClosed) {
                                        return null;
                                    }
                                }
                                return null;
                            }
                        });
                if (submissionResult.startsWith("alert:")) {
                    String alertText = submissionResult.substring("alert:".length()).trim();
                    driver.switchTo().alert().accept();
                    throw new IllegalStateException(
                            "The site rejected the Check-In submission: " + alertText
                                    + ". Verify that Brought in By and Previous Owner match people "
                                    + "already in the site's database.");
                }
            } catch (TimeoutException e) {
                Object validationIssues = ((JavascriptExecutor) driver).executeScript(
                        "var button = document.querySelector(\"input[type='submit'][name='submit'][value='Add']\");" +
                                "var form = button && button.form;" +
                                "if (!form) return 'Check-In form is no longer present';" +
                                "return Array.from(form.querySelectorAll(':invalid')).map(function(el) {" +
                                "  return (el.name || el.id || el.tagName) + ': ' + el.validationMessage;" +
                                "}).join('; ') || 'No browser validation errors reported';");
                printPageSnapshot(driver, "Check-In submission did not complete");
                throw new IllegalStateException(
                        "Check-In was not confirmed after clicking Add. Current URL: "
                                + driver.getCurrentUrl() + ". Form validation: " + validationIssues,
                        e);
            }

            // Row clearing is disabled so the same test data can be reused.
            // deleteRow(checkinSheet, 1);
            return true;
        }
        return false;
    }

    private static void addVaccinations(WebDriver driver, Sheet vaccSheet) {
        // Implementation for adding vaccinations based on the data in vaccSheet
        String catName = (vaccSheet.getRow(1) != null && vaccSheet.getRow(1).getCell(0) != null) ? getCellValueAsString(vaccSheet.getRow(1).getCell(0)) : "";
        String neededOrGiven = (vaccSheet.getRow(1) != null && vaccSheet.getRow(1).getCell(1) != null) ? getCellValueAsString(vaccSheet.getRow(1).getCell(1)) : "";
        LocalDate vaccDate = (vaccSheet.getRow(1) != null && vaccSheet.getRow(1).getCell(2) != null) ? getCellValueAsDate(vaccSheet.getRow(1).getCell(2)) : null;
        String vaccType = (vaccSheet.getRow(1) != null && vaccSheet.getRow(1).getCell(3) != null) ? getCellValueAsString(vaccSheet.getRow(1).getCell(3)) : "";
        String vet = (vaccSheet.getRow(1) != null && vaccSheet.getRow(1).getCell(4) != null) ? getCellValueAsString(vaccSheet.getRow(1).getCell(4)) : "";
        String vetClinic = (vaccSheet.getRow(1) != null && vaccSheet.getRow(1).getCell(5) != null) ? getCellValueAsString(vaccSheet.getRow(1).getCell(5)) : "";
        String genComments = (vaccSheet.getRow(1) != null && vaccSheet.getRow(1).getCell(6) != null) ? getCellValueAsString(vaccSheet.getRow(1).getCell(6)) : "";
        String hidComments = (vaccSheet.getRow(1) != null && vaccSheet.getRow(1).getCell(7) != null) ? getCellValueAsString(vaccSheet.getRow(1).getCell(7)) : "";

        if (hasData(vaccSheet)) {
            openAnimalFromSearch(driver, catName, "Vaccinations!A2");
            driver.findElement(By.name("history")).click();
            driver.findElement(By.xpath("//option[text()='Medical History']")).click();
            driver.findElement(By.cssSelector("[title='Add a Vaccination Record']")).click();
            
            //Determine if the date given is for an upcoming vaccination or one already received
            if ("Needed".equals(neededOrGiven)) {
                driver.findElement(By.name("nd")).click();
                driver.findElement(By.name("nd")).sendKeys(Keys.chord(Keys.CONTROL + "a"), Keys.DELETE);
                driver.findElement(By.name("nd")).sendKeys(vaccDate.toString());
                driver.findElement(By.name("nd")).sendKeys(Keys.TAB);
            } else if ("Given".equals(neededOrGiven)) {
                driver.findElement(By.name("gd")).click();
                driver.findElement(By.name("gd")).sendKeys(Keys.chord(Keys.CONTROL + "a"), Keys.DELETE);
                driver.findElement(By.name("gd")).sendKeys(vaccDate.toString());
                driver.findElement(By.name("gd")).sendKeys(Keys.TAB);
            }

            if (vaccType == "Feline Panleukopenia") {
                driver.findElement(By.cssSelector("input[value='26569']")).click();
            } else if (vaccType == "Rabies") {
                driver.findElement(By.cssSelector("input[value='26574']")).click();
            } else if (vaccType == "Rhino Trachitis/Feline Calicivirus") {
                driver.findElement(By.cssSelector("input[value='26571']")).click();
            }

            driver.findElement(By.name("v")).sendKeys(vet);
            driver.findElement(By.name("vc")).sendKeys(vetClinic);
            driver.findElement(By.name("gc")).sendKeys(genComments);
            driver.findElement(By.name("hc")).sendKeys(hidComments);
            driver.findElement(By.cssSelector("input[value='Add']")).click();

            // Row clearing is disabled so the same test data can be reused.
            // deleteRow(vaccSheet, 1);
        }
    }

    private static void addMedCheck(WebDriver driver, Sheet medCheckSheet) {
        // Implementation for adding a medical check-up based on the data in medCheckSheet        
        String nameofCat = (medCheckSheet.getRow(1) != null && medCheckSheet.getRow(1).getCell(0) != null) ? getCellValueAsString(medCheckSheet.getRow(1).getCell(0)) : "";
        String neededOrGiven = (medCheckSheet.getRow(1) != null && medCheckSheet.getRow(1).getCell(1) != null) ? getCellValueAsString(medCheckSheet.getRow(1).getCell(1)) : "";
        LocalDate examDate = (medCheckSheet.getRow(1) != null && medCheckSheet.getRow(1).getCell(2) != null) ? getCellValueAsDate(medCheckSheet.getRow(1).getCell(2)) : null;
        String vet = (medCheckSheet.getRow(1) != null && medCheckSheet.getRow(1).getCell(3) != null) ? getCellValueAsString(medCheckSheet.getRow(1).getCell(3)) : "";
        String vetClinic = (medCheckSheet.getRow(1) != null && medCheckSheet.getRow(1).getCell(4) != null) ? getCellValueAsString(medCheckSheet.getRow(1).getCell(4)) : "";
        String genComments = (medCheckSheet.getRow(1) != null && medCheckSheet.getRow(1).getCell(5) != null) ? getCellValueAsString(medCheckSheet.getRow(1).getCell(5)) : "";
        String hidComments = (medCheckSheet.getRow(1) != null && medCheckSheet.getRow(1).getCell(6) != null) ? getCellValueAsString(medCheckSheet.getRow(1).getCell(6)) : "";
        String recommendations = (medCheckSheet.getRow(1) != null && medCheckSheet.getRow(1).getCell(7) != null) ? getCellValueAsString(medCheckSheet.getRow(1).getCell(7)) : "";
        String weight = (medCheckSheet.getRow(1) != null && medCheckSheet.getRow(1).getCell(8) != null) ? getCellValueAsString(medCheckSheet.getRow(1).getCell(8)) : "";
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));

        if (hasData(medCheckSheet)) {
            driver.findElement(By.id("s")).sendKeys(nameofCat);
            driver.findElement(By.id("submit")).click();
            driver.findElement(By.xpath("//a[@class='animal'][text()='" + nameofCat + "']")).click();
            driver.findElement(By.name("history")).click();
            wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//option[text()='Medical History']"))).click();
            driver.findElement(By.cssSelector("[title='Add a Medical Check-Up']")).click();

            //Determine if the date given is for an upcoming vaccination or one already received
            if ("Needed".equals(neededOrGiven)) {
                driver.findElement(By.name("nd")).click();
                driver.findElement(By.name("nd")).sendKeys(Keys.chord(Keys.CONTROL + "a"), Keys.DELETE);
                driver.findElement(By.name("nd")).sendKeys(examDate.toString());
                driver.findElement(By.name("nd")).sendKeys(Keys.TAB);
            } else if ("Given".equals(neededOrGiven)) {
                driver.findElement(By.name("gd")).click();
                driver.findElement(By.name("gd")).sendKeys(Keys.chord(Keys.CONTROL + "a"), Keys.DELETE);
                driver.findElement(By.name("gd")).sendKeys(examDate.toString());
                driver.findElement(By.name("gd")).sendKeys(Keys.TAB);
            }
            driver.findElement(By.name("v")).sendKeys(vet);
            driver.findElement(By.name("vc")).sendKeys(vetClinic);
            driver.findElement(By.name("gc")).sendKeys(genComments);
            driver.findElement(By.name("hc")).sendKeys(hidComments);
            driver.findElement(By.name("r")).sendKeys(recommendations);        
            driver.findElement(By.name("w")).sendKeys(weight);
            driver.findElement(By.cssSelector("input[value='Add']")).click();
        
            // Row clearing is disabled so the same test data can be reused.
            // deleteRow(medCheckSheet, 1);
        }
    }

    private static void addProcedure(WebDriver driver, Sheet procedureSheet) {
        // Implementation for adding a medical procedure based on the data in procedureSheet
        String nameofCat = (procedureSheet.getRow(1) != null && procedureSheet.getRow(1).getCell(0) != null) ? getCellValueAsString(procedureSheet.getRow(1).getCell(0)) : "";
        String neededOrGiven = (procedureSheet.getRow(1) != null && procedureSheet.getRow(1).getCell(1) != null) ? getCellValueAsString(procedureSheet.getRow(1).getCell(1)) : "";
        LocalDate examDate = (procedureSheet.getRow(1) != null && procedureSheet.getRow(1).getCell(2) != null) ? getCellValueAsDate(procedureSheet.getRow(1).getCell(2)) : null;
        String procedureType = (procedureSheet.getRow(1) != null && procedureSheet.getRow(1).getCell(3) != null) ? getCellValueAsString(procedureSheet.getRow(1).getCell(3)) : "";
        String vet = (procedureSheet.getRow(1) != null && procedureSheet.getRow(1).getCell(4) != null) ? getCellValueAsString(procedureSheet.getRow(1).getCell(4)) : "";
        String vetClinic = (procedureSheet.getRow(1) != null && procedureSheet.getRow(1).getCell(5) != null) ? getCellValueAsString(procedureSheet.getRow(1).getCell(5)) : "";
        String genComments = (procedureSheet.getRow(1) != null && procedureSheet.getRow(1).getCell(6) != null) ? getCellValueAsString(procedureSheet.getRow(1).getCell(6)) : "";
        String hidComments = (procedureSheet.getRow(1) != null && procedureSheet.getRow(1).getCell(7) != null) ? getCellValueAsString(procedureSheet.getRow(1).getCell(7)) : "";
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));

        if (hasData(procedureSheet)) {
            driver.findElement(By.id("s")).sendKeys(nameofCat);
            driver.findElement(By.id("submit")).click();
            driver.findElement(By.xpath("//a[@class='animal'][text()='" + nameofCat + "']")).click();
            driver.findElement(By.name("history")).click();
            wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//option[text()='Medical History']"))).click();
            driver.findElement(By.cssSelector("[title='Add a Medical Check-Up']")).click();

            //Determine if the date given is for an upcoming vaccination or one already received
            if ("Needed".equals(neededOrGiven)) {
                driver.findElement(By.name("nd")).click();
                driver.findElement(By.name("nd")).sendKeys(Keys.chord(Keys.CONTROL + "a"), Keys.DELETE);
                driver.findElement(By.name("nd")).sendKeys(examDate.toString());
                driver.findElement(By.name("nd")).sendKeys(Keys.TAB);
            } else if ("Given".equals(neededOrGiven)) {
                driver.findElement(By.name("gd")).click();
                driver.findElement(By.name("gd")).sendKeys(Keys.chord(Keys.CONTROL + "a"), Keys.DELETE);
                driver.findElement(By.name("gd")).sendKeys(examDate.toString());
                driver.findElement(By.name("gd")).sendKeys(Keys.TAB);
            }
            driver.findElement(By.name("v")).sendKeys(vet);
            driver.findElement(By.name("vc")).sendKeys(vetClinic);
            driver.findElement(By.name("gc")).sendKeys(genComments);
            driver.findElement(By.name("hc")).sendKeys(hidComments);
            driver.findElement(By.cssSelector("input[value='Add']")).click();

            // Row clearing is disabled so the same test data can be reused.
            // deleteRow(procedureSheet, 1);
        }    
    }

    private static void addPrescriptions(WebDriver driver, Sheet prescriptionSheet) {
        // Implementation for adding prescriptions based on the data in prescriptionSheet
        String nameofCat = (prescriptionSheet.getRow(1) != null && prescriptionSheet.getRow(1).getCell(0) != null) ? getCellValueAsString(prescriptionSheet.getRow(1).getCell(0)) : "";
        String medicationName = (prescriptionSheet.getRow(1) != null && prescriptionSheet.getRow(1).getCell(1) != null) ? getCellValueAsString(prescriptionSheet.getRow(1).getCell(1)) : "";
        String route = (prescriptionSheet.getRow(1) != null && prescriptionSheet.getRow(1).getCell(2) != null) ? getCellValueAsString(prescriptionSheet.getRow(1).getCell(2)) : "";
        String dosage = (prescriptionSheet.getRow(1) != null && prescriptionSheet.getRow(1).getCell(3) != null) ? getCellValueAsString(prescriptionSheet.getRow(1).getCell(3)) : "";
        String doseType = (prescriptionSheet.getRow(1) != null && prescriptionSheet.getRow(1).getCell(4) != null) ? getCellValueAsString(prescriptionSheet.getRow(1).getCell(4)) : "";
        LocalDate oneTimeDateNeeded = (prescriptionSheet.getRow(1) != null && prescriptionSheet.getRow(1).getCell(5) != null) ? getCellValueAsDate(prescriptionSheet.getRow(1).getCell(5)) : null;
        LocalDate oneTimeDateGiven = (prescriptionSheet.getRow(1) != null && prescriptionSheet.getRow(1).getCell(6) != null) ? getCellValueAsDate(prescriptionSheet.getRow(1).getCell(6)) : null;
        String administerer = (prescriptionSheet.getRow(1) != null && prescriptionSheet.getRow(1).getCell(7) != null) ? getCellValueAsString(prescriptionSheet.getRow(1).getCell(7)) : "";
        String createMAR = (prescriptionSheet.getRow(1) != null && prescriptionSheet.getRow(1).getCell(8) != null) ? getCellValueAsString(prescriptionSheet.getRow(1).getCell(8)) : "";
        String perDay = (prescriptionSheet.getRow(1) != null && prescriptionSheet.getRow(1).getCell(9) != null) ? getCellValueAsString(prescriptionSheet.getRow(1).getCell(9)) : "";
        String dateRange = (prescriptionSheet.getRow(1) != null && prescriptionSheet.getRow(1).getCell(10) != null) ? getCellValueAsString(prescriptionSheet.getRow(1).getCell(10)) : "";
        LocalDate recurringDateStart = (prescriptionSheet.getRow(1) != null && prescriptionSheet.getRow(1).getCell(11) != null) ? getCellValueAsDate(prescriptionSheet.getRow(1).getCell(11)) : null;
        LocalDate recurringDateEnd = (prescriptionSheet.getRow(1) != null && prescriptionSheet.getRow(1).getCell(12) != null) ? getCellValueAsDate(prescriptionSheet.getRow(1).getCell(12)) : null;
        String prescriber = (prescriptionSheet.getRow(1) != null && prescriptionSheet.getRow(1).getCell(13) != null) ? getCellValueAsString(prescriptionSheet.getRow(1).getCell(13)) : "";
        String vetClinic = (prescriptionSheet.getRow(1) != null && prescriptionSheet.getRow(1).getCell(14) != null) ? getCellValueAsString(prescriptionSheet.getRow(1).getCell(14)) : "";
        String genComments = (prescriptionSheet.getRow(1) != null && prescriptionSheet.getRow(1).getCell(15) != null) ? getCellValueAsString(prescriptionSheet.getRow(1).getCell(15)) : "";
        String hidComments = (prescriptionSheet.getRow(1) != null && prescriptionSheet.getRow(1).getCell(16) != null) ? getCellValueAsString(prescriptionSheet.getRow(1).getCell(16)) : "";
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));

        if (hasData(prescriptionSheet)) {
            driver.findElement(By.id("s")).sendKeys(nameofCat);
            driver.findElement(By.id("submit")).click();
            driver.findElement(By.xpath("//a[@class='animal'][text()='" + nameofCat + "']")).click();
            driver.findElement(By.name("history")).click();
            wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//option[text()='Medication Prescriptions']"))).click();
            driver.findElement(By.cssSelector("[title='Add a Medication Prescription']")).click();

            driver.findElement(By.name("mn")).sendKeys(medicationName);
            driver.findElement(By.name("r")).sendKeys(route);
            driver.findElement(By.name("d")).sendKeys(dosage);
            
            if ("One Time".equals(doseType)) {
                driver.findElement(By.cssSelector("input[value='ot']")).click();
            } else if ("TimeRange".equals(doseType)) {
                driver.findElement(By.cssSelector("input[value='tb']")).click();
            }

            //If "One Time" is selected in the previous step, otherwise refer to "recurringDateStart" and "recurringDateEnd" for the date inputs
            if (oneTimeDateNeeded != null) {
                driver.findElement(By.name("otdn")).sendKeys(Keys.chord(Keys.CONTROL + "a"), Keys.DELETE);
                driver.findElement(By.name("otdn")).sendKeys(oneTimeDateNeeded.toString());
                driver.findElement(By.name("otdn")).sendKeys(Keys.TAB);
            }
            if (oneTimeDateGiven != null) {
                driver.findElement(By.name("otdg")).sendKeys(Keys.chord(Keys.CONTROL + "a"), Keys.DELETE);
                driver.findElement(By.name("otdg")).sendKeys(oneTimeDateGiven.toString());
                driver.findElement(By.name("otdg")).sendKeys(Keys.TAB);
            }
            driver.findElement(By.name("otp")).sendKeys(administerer);
            
            //Create a medication administration record (MAR) if "Yes" is selected, otherwise skip to filling out the prescription details and just have it linked to the cat's medical history without a MAR entry
            if ("Yes".equals(createMAR)) {
                driver.findElement(By.name("tbcm")).click();
            }
            driver.findElement(By.name("tbdf")).sendKeys(perDay);
            driver.findElement(By.name("tbgf")).sendKeys(dateRange);

            //See section above for "One Time"
            if (recurringDateStart != null) {
                driver.findElement(By.name("tbsd")).click();
                driver.findElement(By.name("tbsd")).sendKeys(Keys.chord(Keys.CONTROL + "a"), Keys.DELETE);
                driver.findElement(By.name("tbsd")).sendKeys(recurringDateStart.toString());
                driver.findElement(By.name("tbsd")).sendKeys(Keys.TAB);
            }
            if (recurringDateEnd != null) {
                driver.findElement(By.name("tbed")).click();
                driver.findElement(By.name("tbed")).sendKeys(Keys.chord(Keys.CONTROL + "a"), Keys.DELETE);
                driver.findElement(By.name("tbed")).sendKeys(recurringDateEnd.toString());
                driver.findElement(By.name("tbed")).sendKeys(Keys.TAB);
            }

            //Below information is for either "One Time" or "Time Range" prescriptions, so fill it out regardless of which dose type is selected
            driver.findElement(By.name("p")).sendKeys(prescriber);
            driver.findElement(By.name("vc")).sendKeys(vetClinic);
            driver.findElement(By.name("gc")).sendKeys(genComments);
            driver.findElement(By.name("hc")).sendKeys(hidComments);
            driver.findElement(By.cssSelector("input[value='Add']")).click();

            // Row clearing is disabled so the same test data can be reused.
            // deleteRow(prescriptionSheet, 1);
        }
    }

    private static LocalDate getCellValueAsDate(Cell dateCell) {
        if (dateCell == null || dateCell.getCellType() != CellType.NUMERIC || !DateUtil.isCellDateFormatted(dateCell)) {
            return null;
        }
        Date date = dateCell.getDateCellValue();
        return date.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
    }

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
}
