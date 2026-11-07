package demo;

import helpers.DateHelper;
import helpers.FormHelper;
import org.openqa.selenium.By;
import org.openqa.selenium.Alert;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.NoAlertPresentException;
import org.openqa.selenium.TimeoutException;
import org.openqa.selenium.UnhandledAlertException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import org.apache.poi.ss.usermodel.*;

import java.time.Duration;
import java.time.LocalDate;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;


import static demo.WorkflowSupport.*;

final class CheckInWorkflow {
    static void createCheckInRecord(
            WebDriver driver, Sheet checkinSheet, int rowIndex, String newlyAddedAnimalName) {
        int spreadsheetRow = rowIndex + 1;
        String numberCats = getCellValueAsString(getCell(checkinSheet, rowIndex, 0));
        int numCats;
        try {
            numCats = Integer.parseInt(numberCats.trim());
        } catch (NumberFormatException e) {
            throw new IllegalArgumentException(
                    "New Intake - Check-In!A" + spreadsheetRow
                            + " must contain a whole-number cat count.", e);
        }
        String catNames = getCellValueAsString(getCell(checkinSheet, rowIndex, 1));
        List<String> catList = Arrays.stream(catNames.split(";"))
                .map(String::trim)
                .filter(name -> !name.isEmpty())
                .collect(Collectors.toList());
        if (numCats <= 0 || catList.size() != numCats) {
            throw new IllegalArgumentException(
                    "New Intake - Check-In row " + spreadsheetRow + " lists " + numCats
                            + " cats but contains " + catList.size() + " non-empty names.");
        }
        String broughtBy = getCellValueAsString(getCell(checkinSheet, rowIndex, 2));
        String prevOwner = getCellValueAsString(getCell(checkinSheet, rowIndex, 3));
        String prevAgency = getCellValueAsString(getCell(checkinSheet, rowIndex, 4));
        String surrReason = getCellValueAsString(getCell(checkinSheet, rowIndex, 5));
        String checkinType = getCellValueAsString(getCell(checkinSheet, rowIndex, 6));
        Cell checkinDateCell = getCell(checkinSheet, rowIndex, 7);
        LocalDate checkinDate = checkinDateCell == null ? null : DateHelper.getCellValueAsDate(checkinDateCell);
        String checker = getCellValueAsString(getCell(checkinSheet, rowIndex, 8));
        String sheltAnCounts = getCellValueAsString(getCell(checkinSheet, rowIndex, 9));
        String genComments = getCellValueAsString(getCell(checkinSheet, rowIndex, 10));
        String hidComments = getCellValueAsString(getCell(checkinSheet, rowIndex, 11));
        String move = getCellValueAsString(getCell(checkinSheet, rowIndex, 12));
        String currentlyWith = getCellValueAsString(getCell(checkinSheet, rowIndex, 13));
        String location = getCellValueAsString(getCell(checkinSheet, rowIndex, 14));
        String subLocation = getCellValueAsString(getCell(checkinSheet, rowIndex, 15));
        LocalDate today = LocalDate.now();

        if (hasCheckInCatNames(checkinSheet, rowIndex)) {
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
                openAnimalFromSearch(
                        driver, catList.get(0), "New Intake - Check-In!B" + spreadsheetRow);
                openAnimalHistoryAddForm(
                        driver, catList.get(0), "Check-In History", "check.php", "Add a Check-In").click();
            }

            WebElement firstAnimalField = new WebDriverWait(driver, Duration.ofSeconds(10))
                    .until(ExpectedConditions.visibilityOfElementLocated(By.name("aa1")));
            String selectedFirstAnimal = firstAnimalField.getDomProperty("value");
            if (!catList.get(0).equals(selectedFirstAnimal == null ? "" : selectedFirstAnimal.trim())) {
                throw new IllegalStateException(
                        "The Check-In form's first animal is '" + selectedFirstAnimal
                                + "', but the first name in New Intake - Check-In!B"
                                + spreadsheetRow + " is '"
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

            By addButton = By.cssSelector(
                    "input[type='submit'][id='submit'][name='submit'][value='Add']");
            String formUrl = driver.getCurrentUrl();
            try {
                Thread.sleep(1_000);
            } catch (InterruptedException e) {
                Thread.currentThread().interrupt();
                throw new IllegalStateException("Interrupted while pausing before Check-In submission.", e);
            }
            new WebDriverWait(driver, Duration.ofSeconds(10))
                    .until(ExpectedConditions.elementToBeClickable(addButton))
                    .click();
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
                        "var button = document.querySelector(\"input[type='submit'][id='submit'][name='submit'][value='Add']\");" +
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
        }

    }
}
