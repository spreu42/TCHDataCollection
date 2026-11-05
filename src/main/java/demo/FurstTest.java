package demo;

import helpers.DateHelper;
import helpers.AutocompleteHelper;
import helpers.FormHelper;
import org.openqa.selenium.By;
import org.openqa.selenium.Alert;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.NoAlertPresentException;
import org.openqa.selenium.StaleElementReferenceException;
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
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
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
                if (hasMedicalRecordData(vaccSheet, "Vaccinations")) {
                    addVaccinations(driver, vaccSheet);
                }
                if (hasMedicalRecordData(medCheckSheet, "Medical Check-Up")) {
                    addMedCheck(driver, medCheckSheet);
                }
                if (hasMedicalRecordData(procedureSheet, "Medical Procedure")) {
                    addProcedure(driver, procedureSheet);
                }
                if (hasPrescriptionData(prescriptionSheet)) {
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

    private static boolean hasMedicalRecordData(Sheet sheet, String sheetName) {
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

    private static boolean hasPrescriptionData(Sheet prescriptionSheet) {
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

    private static void printPageSnapshot(WebDriver driver, String label) {
        Object text = ((JavascriptExecutor) driver).executeScript(
                "var c = document.getElementById('content') || document.body;" +
                        "return c.innerText.substring(0, 5000);");
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
        AutocompleteHelper.selectMatchingSuggestion(
                driver, suggestions, normalizedName, fieldLabel + " from the site's suggestions");
    }

    private static void openAnimalFromSearch(WebDriver driver, String animalName, String sourceCell) {
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

    private static WebElement openAnimalHistoryAddForm(
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

    private static String getCellValueAsString(Cell cell) {
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
        LocalDate statusChangeDate = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(9) != null) ? DateHelper.getCellValueAsDate(newCatSheet.getRow(1).getCell(9)) : null;
        String gender = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(10) != null) ? getCellValueAsString(newCatSheet.getRow(1).getCell(10)) : "";
        String catColor = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(11) != null) ? getCellValueAsString(newCatSheet.getRow(1).getCell(11)) : "";
        String chipType = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(12) != null) ? getCellValueAsString(newCatSheet.getRow(1).getCell(12)) : "";
        String chipNumber = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(13) != null) ? getCellValueAsString(newCatSheet.getRow(1).getCell(13)) : "";
        String tagNumber = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(14) != null) ? getCellValueAsString(newCatSheet.getRow(1).getCell(14)) : "";
        String DOBest = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(15) != null) ? getCellValueAsString(newCatSheet.getRow(1).getCell(15)) : "";
        LocalDate DOB = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(16) != null) ? DateHelper.getCellValueAsDate(newCatSheet.getRow(1).getCell(16)) :null;
        LocalDate altered = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(17) != null) ? DateHelper.getCellValueAsDate(newCatSheet.getRow(1).getCell(17)) : null;
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
                FormHelper.setCheckbox(driver, "rta", true);
            }
            if ("Yes".equals(webDisplay)) {
                FormHelper.setCheckbox(driver, "sw", true);
            }
            driver.findElement(By.name("st")).sendKeys(catStatus);
            if (statusChangeDate != null) {
                DateHelper.enterDateText(driver, "std", statusChangeDate, false);
            }
            driver.findElement(By.name("g")).sendKeys(gender);
            driver.findElement(By.name("pc")).sendKeys(catColor);
            driver.findElement(By.name("mt")).sendKeys(chipType);
            driver.findElement(By.name("mn")).sendKeys(chipNumber);
            driver.findElement(By.name("tn")).sendKeys(tagNumber);
            
            if (DOB != null) {
                WebElement dobField = driver.findElement(By.name("bd"));
                DateHelper.selectDateFromPicker(driver, dobField, DOB);
            }

            if (altered != null) {
                DateHelper.enterDateText(driver, "nd", altered, true);
            }

            FormHelper.enterText(driver, "gc", genComments);
            FormHelper.enterText(driver, "hc", hidCommments);
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
        LocalDate checkinDate = checkinDateCell == null ? null : DateHelper.getCellValueAsDate(checkinDateCell);
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
                openAnimalHistoryAddForm(
                        driver, catList.get(0), "Check-In History", "check.php", "Add a Check-In").click();
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
            FormHelper.enterText(driver, "j", prevAgency);
            logCheckInField(driver, By.name("j"), "j (previous agency)", prevAgency);
            FormHelper.enterText(driver, "sr", surrReason);
            logCheckInField(driver, By.name("sr"), "sr (surrender reason)", surrReason);
            FormHelper.enterText(driver, "ct", checkinType);
            logCheckInField(driver, By.name("ct"), "ct (check-in type)", checkinType);
            LocalDate submittedCheckinDate = checkinDate == null ? today : checkinDate;
            DateHelper.enterDateText(driver, "cd", submittedCheckinDate, true);
            logCheckInField(driver, By.name("cd"), "cd (check-in date)", submittedCheckinDate.toString());
            FormHelper.enterText(driver, "cb", checker);
            logCheckInField(driver, By.name("cb"), "cb (checker)", checker);
            FormHelper.enterText(driver, "sacs", sheltAnCounts);
            logCheckInField(driver, By.name("sacs"), "sacs (shelter animal counts)", sheltAnCounts);
            FormHelper.enterText(driver, "gc", genComments);
            logCheckInField(driver, By.name("gc"), "gc (general comments)", genComments);
            FormHelper.enterText(driver, "hc", hidComments);
            logCheckInField(driver, By.name("hc"), "hc (hidden comments)", hidComments);
            
            //Process a move
            By moveCheckbox = By.name("mrc");
            if ("Yes".equals(move)) {
                FormHelper.setCheckbox(driver, "mrc", true);
                logCheckInField(driver, moveCheckbox, "mrc (move requested)", move);
                selectPersonFromAutocomplete(
                        driver, "mrcw", "mrcwa", "mrp", currentlyWith, "Currently With");
                logCheckInField(driver, By.name("mrcw"), "mrcw (currently with)", currentlyWith);
                FormHelper.enterText(driver, "mrl", location);
                logCheckInField(driver, By.name("mrl"), "mrl (location)", location);
                FormHelper.enterText(driver, "mrsl", subLocation);
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

    private static void addMedCheck(WebDriver driver, Sheet medCheckSheet) {
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

    private static void addProcedure(WebDriver driver, Sheet procedureSheet) {
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

    private static void addPrescriptions(WebDriver driver, Sheet prescriptionSheet) {
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

    private static void addPrescriptionRecord(WebDriver driver, Sheet prescriptionSheet, int rowIndex,
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

    private static void selectPrescriptionMedication(WebDriver driver, String medicationName, int spreadsheetRow) {
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

    private static void selectNumericPrescriptionOption(
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
