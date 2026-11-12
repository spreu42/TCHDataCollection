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
import org.openqa.selenium.WebDriverException;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import org.apache.poi.ss.usermodel.*;

import java.io.IOException;
import java.nio.file.Path;
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
        String littermatesValue = getCellValueAsString(getCell(checkinSheet, rowIndex, 16)).trim();
        String motherName = getCellValueAsString(getCell(checkinSheet, rowIndex, 17)).trim();
        String littermatesHeader = getCellValueAsString(getCell(checkinSheet, 0, 16)).trim();
        boolean addLitter = LitterWorkflow.isRequested(littermatesValue, spreadsheetRow);
        System.out.printf("[Check-In] row %d: column Q ('%s')='%s'; litter=%s%n",
                spreadsheetRow,
                littermatesHeader,
                littermatesValue,
                addLitter ? "add" : "skip");
        if (addLitter) {
            LitterWorkflow.validateRequest(catList, motherName, spreadsheetRow);
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

        if (addLitter) {
            LitterWorkflow.addLitter(driver, catList, motherName, spreadsheetRow);
        }

        if (hasCheckInCatNames(checkinSheet, rowIndex)) {
            if (!addLitter
                    && newlyAddedAnimalName != null
                    && newlyAddedAnimalName.equals(catList.get(0))) {
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

            new WebDriverWait(driver, Duration.ofSeconds(10))
                    .until(ExpectedConditions.visibilityOfElementLocated(By.name("aa1")));
            selectPersonFromAutocomplete(
                    driver, "aa1", "aaa1", "a1", catList.get(0), "Check-In animal 1");
            WebElement firstAnimalField = driver.findElement(By.name("aa1"));
            String selectedFirstAnimal = firstAnimalField.getDomProperty("value");
            if (!catList.get(0).equals(selectedFirstAnimal == null ? "" : selectedFirstAnimal.trim())) {
                throw new IllegalStateException(
                        "The Check-In form's first animal is '" + selectedFirstAnimal
                                + "', but the first name in New Intake - Check-In!B"
                                + spreadsheetRow + " is '"
                                + catList.get(0) + "'.");
            }
            logCheckInField(driver, By.name("aa1"), "first animal", catList.get(0));
            logCheckInField(driver, By.name("a1"), "first animal linked record", catList.get(0));

            // The first animal is selected above; add and select only the remaining names.
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
                logCheckInField(driver, By.name("a" + fieldNumber),
                        "additional animal " + fieldNumber + " linked record", catList.get(i));
            }

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
            if ("Yes".equalsIgnoreCase(move.trim()) || "Y".equalsIgnoreCase(move.trim())) {
                FormHelper.setCheckbox(driver, "mrc", true);
                logCheckInField(driver, moveCheckbox, "mrc (move requested)", move);
                selectPersonFromAutocomplete(
                        driver, "mrcw", "mrcwa", "mrp", currentlyWith, "Currently With");
                logCheckInField(driver, By.name("mrcw"), "mrcw (currently with)", currentlyWith);
                logCheckInField(driver, By.name("mrp"), "currently-with linked record", currentlyWith);
                FormHelper.enterText(driver, "mrl", location);
                logCheckInField(driver, By.name("mrl"), "mrl (location)", location);
                FormHelper.enterText(driver, "mrsl", subLocation);
                logCheckInField(driver, By.name("mrsl"), "mrsl (sub-location)", subLocation);
            } else {
                logCheckInField(driver, moveCheckbox, "mrc (move requested)", move);
            }

            selectPersonFromAutocomplete(driver, "bb", "bba", "b", broughtBy, "Brought in By");
            logCheckInField(driver, By.name("bb"), "bb (brought by)", broughtBy);
            logCheckInField(driver, By.name("b"), "brought-by linked record", broughtBy);
            selectPersonFromAutocomplete(driver, "po", "poa", "p", prevOwner, "Previous Owner");
            logCheckInField(driver, By.name("po"), "po (previous owner)", prevOwner);
            logCheckInField(driver, By.name("p"), "previous-owner linked record", prevOwner);

            By addButton = By.cssSelector(
                    "input[type='submit'][id='submit'][name='submit'][value='Add']");
            String formUrl = driver.getCurrentUrl();
            try {
                Thread.sleep(1_000);
            } catch (InterruptedException e) {
                Thread.currentThread().interrupt();
                throw new IllegalStateException("Interrupted while pausing before Check-In submission.", e);
            }
            WebElement addSubmitButton = new WebDriverWait(driver, Duration.ofSeconds(10))
                    .until(ExpectedConditions.elementToBeClickable(addButton));
            logCheckInSubmissionFields(driver, broughtBy, prevOwner, spreadsheetRow);
            addSubmitButton.click();
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
                    IllegalStateException failure = new IllegalStateException(
                            "The site rejected the Check-In submission: " + alertText
                                    + ". The check-in was not confirmed.");
                    try {
                        printPageSnapshot(driver, "Check-In submission rejected");
                        Path screenshotPath = captureScreenshot(
                                driver, "check-in-row-" + spreadsheetRow);
                        System.err.println("[Workflow failure] Screenshot saved to: "
                                + screenshotPath.toAbsolutePath());
                    } catch (IOException | WebDriverException diagnosticFailure) {
                        failure.addSuppressed(diagnosticFailure);
                        System.err.println("[Workflow failure] Could not capture page diagnostics: "
                                + diagnosticFailure.getMessage());
                    }
                    throw failure;
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

    private static void logCheckInSubmissionFields(
            WebDriver driver, String broughtBy, String previousOwner, int spreadsheetRow) {
        String broughtBySyncResult = synchronizeCheckInPersonLink(driver, "bb", "b", broughtBy);
        System.out.println("[Check-In debug] brought-by form link: " + broughtBySyncResult);
        String previousOwnerSyncResult = synchronizeCheckInPersonLink(driver, "po", "p", previousOwner);
        System.out.println("[Check-In debug] previous-owner form link: " + previousOwnerSyncResult);

        Object submittedFields = ((JavascriptExecutor) driver).executeScript(
                "var button = document.querySelector(arguments[0]);" +
                        "var form = button && button.form;" +
                        "if (!form) return 'Add button is not attached to a form';" +
                        "var relevant = /^(aa\\d+|a\\d+|bb|b|po|p|mrc|mrcw|mrp|mrl|mrsl)$/;" +
                        "return JSON.stringify(Array.from(new FormData(form).entries())" +
                        "  .filter(function(entry) { return relevant.test(entry[0]); }));",
                "input[type='submit'][id='submit'][name='submit'][value='Add']");
        System.out.println("[Check-In debug] submitted form animal/person/move fields="
                + submittedFields);
        String submittedBroughtBy = getCheckInSubmissionField(driver, "b");
        if (!broughtBy.trim().isEmpty() && submittedBroughtBy.isEmpty()) {
            throw new IllegalStateException(
                    "Brought In By in New Intake - Check-In row " + spreadsheetRow
                            + " was selected on screen but is missing from the submitted form. "
                            + "Check-In was not submitted.");
        }
        String submittedPreviousOwner = getCheckInSubmissionField(driver, "p");
        if (!previousOwner.trim().isEmpty() && submittedPreviousOwner.isEmpty()) {
            throw new IllegalStateException(
                    "Previous Owner in New Intake - Check-In row " + spreadsheetRow
                            + " was selected on screen but is missing from the submitted form. "
                            + "Check-In was not submitted.");
        }
    }

    private static String synchronizeCheckInPersonLink(
            WebDriver driver, String textFieldName, String linkedFieldName, String personName) {
        return (String) ((JavascriptExecutor) driver).executeScript(
                "var button = document.querySelector(arguments[0]);" +
                        "var form = button && button.form;" +
                        "if (!form) return 'Add button is not attached to a form';" +
                        "var textFieldName = String(arguments[1]);" +
                        "var linkedFieldName = String(arguments[2]);" +
                        "var personName = String(arguments[3]).trim();" +
                        "if (!personName) return 'person is blank; link sync skipped';" +
                        "var textField = Array.from(form.elements).find(function(el) {" +
                        "  return el.name === textFieldName;" +
                        "});" +
                        "if (!textField || textField.value.trim() !== personName)" +
                        "  return 'selected person text does not match the requested value';" +
                        "var submittedLinks = Array.from(form.elements).filter(function(el) {" +
                        "  return el.name === linkedFieldName;" +
                        "});" +
                        "if (!submittedLinks.length) return 'no linked field is attached to the Add form';" +
                        "if (String(new FormData(form).get(linkedFieldName) || '').trim() !== '')" +
                        "  return 'Add form linked field is already populated';" +
                        "var selector = '[name=\"' + linkedFieldName + '\"]';" +
                        "var candidates = Array.from(document.querySelectorAll(selector))" +
                        "  .map(function(el) { return String(el.value).trim(); })" +
                        "  .filter(function(value) { return value !== ''; });" +
                        "var uniqueCandidates = Array.from(new Set(candidates));" +
                        "if (uniqueCandidates.length !== 1)" +
                        "  return 'could not identify one selected linked ID';" +
                        "submittedLinks.forEach(function(el) {" +
                        "  el.value = uniqueCandidates[0];" +
                        "  el.dispatchEvent(new Event('input', { bubbles: true }));" +
                        "  el.dispatchEvent(new Event('change', { bubbles: true }));" +
                        "});" +
                        "return 'copied selected linked ID into the Add form';",
                "input[type='submit'][id='submit'][name='submit'][value='Add']",
                textFieldName,
                linkedFieldName,
                personName);
    }

    private static String getCheckInSubmissionField(WebDriver driver, String fieldName) {
        return (String) ((JavascriptExecutor) driver).executeScript(
                "var button = document.querySelector(arguments[0]);" +
                        "var form = button && button.form;" +
                        "return form ? (new FormData(form).get(arguments[1]) || '') : '';",
                "input[type='submit'][id='submit'][name='submit'][value='Add']",
                fieldName);
    }
}
