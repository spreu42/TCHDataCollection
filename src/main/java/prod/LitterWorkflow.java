package prod;

import org.openqa.selenium.By;
import org.openqa.selenium.NoSuchFrameException;
import org.openqa.selenium.StaleElementReferenceException;
import org.openqa.selenium.TimeoutException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebDriverException;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.IOException;
import java.nio.file.Path;
import java.time.Duration;
import java.util.List;
import java.util.stream.Collectors;

import static helpers.WorkflowSupport.captureScreenshot;
import static helpers.WorkflowSupport.openAnimalFromSearch;
import static helpers.WorkflowSupport.printPageSnapshot;
import static helpers.WorkflowSupport.selectPersonFromAutocomplete;
import static helpers.WorkflowSupport.waitForFormSubmission;

final class LitterWorkflow {
    private LitterWorkflow() {
    }

    static boolean isRequested(String value, int spreadsheetRow) {
        if ("Yes".equalsIgnoreCase(value)) {
            return true;
        }
        if (value.isEmpty() || "No".equalsIgnoreCase(value)) {
            return false;
        }
        throw new IllegalArgumentException(
                "New Intake - Check-In!Q" + spreadsheetRow
                        + " must be 'Yes', 'No', or blank.");
    }

    static void validateRequest(List<String> catNames, String motherName, int spreadsheetRow) {
        if (motherName.isEmpty()) {
            return;
        }
        long motherMatches = catNames.stream()
                .filter(name -> name.equalsIgnoreCase(motherName))
                .count();
        if (motherMatches > 1) {
            throw new IllegalArgumentException(
                    "New Intake - Check-In!R" + spreadsheetRow
                            + " must not appear more than once in column B when Littermates? is Yes.");
        }
        if (catNames.size() - motherMatches == 0) {
            throw new IllegalArgumentException(
                    "New Intake - Check-In row " + spreadsheetRow
                            + " marks littermates but has no kitten names after excluding the mother, if listed.");
        }
    }

    static void addLitter(
            WebDriver driver, List<String> catNames, String motherName, int spreadsheetRow) {
        String searchName = motherName.isEmpty() ? catNames.get(0) : motherName;
        openAnimalFromSearch(
                driver,
                searchName,
                "New Intake - Check-In!" + (motherName.isEmpty() ? "B" : "R") + spreadsheetRow);

        By litterInformation = By.cssSelector(
                "img[alt='Litter Information'][title='Litter Information'][src='/i/p/stork.gif']");
        new WebDriverWait(driver, Duration.ofSeconds(10))
                .pollingEvery(Duration.ofSeconds(1))
                .until(ExpectedConditions.elementToBeClickable(litterInformation))
                .click();

        By addLitter = By.xpath(
                "//a[starts-with(@href, '/lgp/add.php?a=')"
                        + " and .//img[@src='/i/add.gif'"
                        + " and @title='Add a Litter'"
                        + " and @alt='Add a Litter']]");
        try {
            WebDriverWait litterPageWait = new WebDriverWait(driver, Duration.ofSeconds(10));
            litterPageWait.pollingEvery(Duration.ofSeconds(1));
            litterPageWait.until(currentDriver -> switchToFrameContaining(currentDriver, addLitter));
            WebDriverWait addLitterWait = new WebDriverWait(driver, Duration.ofSeconds(10));
            addLitterWait.pollingEvery(Duration.ofSeconds(1));
            addLitterWait.until(ExpectedConditions.elementToBeClickable(addLitter)).click();

            List<String> kittenNames = catNames.stream()
                    .filter(name -> motherName.isEmpty() || !name.equalsIgnoreCase(motherName))
                    .collect(Collectors.toList());
            if (motherName.isEmpty()) {
                driver.findElement(By.cssSelector("input[type='radio'][name='t'][value='c']")).click();
            }

            By firstKittenField = By.name("aa1");
            new WebDriverWait(driver, Duration.ofSeconds(10))
                    .until(ExpectedConditions.visibilityOfElementLocated(firstKittenField));
            enterKittenName(driver, 1, kittenNames.get(0));

            By addAnotherKitten = By.cssSelector(
                    "a[href='javascript:addElement()'][title='Add an additional animal to keep together']");
            for (int i = 1; i < kittenNames.size(); i++) {
                new WebDriverWait(driver, Duration.ofSeconds(10))
                        .until(ExpectedConditions.elementToBeClickable(addAnotherKitten))
                        .click();
                enterKittenName(driver, i + 1, kittenNames.get(i));
            }

            By submitButton = By.cssSelector(
                    "input[type='submit'][id='submit'][name='submit'][value='Add']");
            String formUrl = driver.getCurrentUrl();
            WebElement submit = new WebDriverWait(driver, Duration.ofSeconds(10))
                    .until(ExpectedConditions.elementToBeClickable(submitButton));
            submit.click();
            waitForFormSubmission(driver, submitButton, formUrl, "Litter", spreadsheetRow);
        } catch (TimeoutException e) {
            IllegalStateException failure = new IllegalStateException(
                    "Timed out waiting for the Litter Information frame or 'Add a Litter' "
                            + "workflow for animal '" + searchName + "' on check-in row "
                            + spreadsheetRow + ". Current page: "
                            + driver.getTitle() + " (" + driver.getCurrentUrl() + ").",
                    e);
            try {
                printPageSnapshot(driver, "Litter workflow timed out");
                Path screenshotPath = captureScreenshot(
                        driver, "litter-add-row-" + spreadsheetRow);
                System.err.println("[Workflow failure] Screenshot saved to: "
                        + screenshotPath.toAbsolutePath());
            } catch (IOException | WebDriverException diagnosticFailure) {
                failure.addSuppressed(diagnosticFailure);
                System.err.println("[Workflow failure] Could not capture page diagnostics: "
                        + diagnosticFailure.getMessage());
            }
            throw failure;
        } finally {
            driver.switchTo().defaultContent();
        }
    }

    private static boolean switchToFrameContaining(WebDriver driver, By target) {
        driver.switchTo().defaultContent();
        if (!driver.findElements(target).isEmpty()) {
            return true;
        }

        List<WebElement> frames = driver.findElements(By.tagName("iframe"));
        for (WebElement frame : frames) {
            try {
                driver.switchTo().frame(frame);
                if (!driver.findElements(target).isEmpty()) {
                    return true;
                }
            } catch (NoSuchFrameException | StaleElementReferenceException ignored) {
                // The profile can replace the frame while its litter panel is loading.
            }
            driver.switchTo().defaultContent();
        }
        return false;
    }

    private static void enterKittenName(WebDriver driver, int fieldNumber, String kittenName) {
        By kittenField = By.name("aa" + fieldNumber);
        new WebDriverWait(driver, Duration.ofSeconds(10))
                .until(ExpectedConditions.visibilityOfElementLocated(kittenField));
        selectPersonFromAutocomplete(
                driver,
                "aa" + fieldNumber,
                "aaa" + fieldNumber,
                "a" + fieldNumber,
                kittenName,
                "Litter kitten " + fieldNumber);
        WebElement field = driver.findElement(kittenField);
        String selectedName = field.getDomProperty("value");
        if (!kittenName.equals(selectedName == null ? "" : selectedName.trim())) {
            throw new IllegalStateException(
                    "Litter kitten field aa" + fieldNumber + " contains '" + selectedName
                            + "' instead of '" + kittenName + "'.");
        }
    }
}
