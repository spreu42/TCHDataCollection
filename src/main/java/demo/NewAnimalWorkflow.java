package demo;

import helpers.DateHelper;
import helpers.FormHelper;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.TimeoutException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import org.apache.poi.ss.usermodel.*;

import java.time.Duration;
import java.time.LocalDate;


import static demo.WorkflowSupport.*;

final class NewAnimalWorkflow {
    static void addNewCat(WebDriver driver, Sheet newCatSheet, int rowIndex) {
        // Implementation for adding a new cat based on the data in newCatSheet
        
        //Gather all filled fields for a new cat
        String nameofCat = (newCatSheet.getRow(rowIndex) != null && newCatSheet.getRow(rowIndex).getCell(0) != null) ? getCellValueAsString(newCatSheet.getRow(rowIndex).getCell(0)) : "";
        String AliasofCat = (newCatSheet.getRow(rowIndex) != null && newCatSheet.getRow(rowIndex).getCell(1) != null) ? getCellValueAsString(newCatSheet.getRow(rowIndex).getCell(1)) : "";
        String species = (newCatSheet.getRow(rowIndex) != null && newCatSheet.getRow(rowIndex).getCell(2) != null) ? getCellValueAsString(newCatSheet.getRow(rowIndex).getCell(2)) : "";
        String breed = (newCatSheet.getRow(rowIndex) != null && newCatSheet.getRow(rowIndex).getCell(3) != null) ? getCellValueAsString(newCatSheet.getRow(rowIndex).getCell(3)) : "";
        String catCode = (newCatSheet.getRow(rowIndex) != null && newCatSheet.getRow(rowIndex).getCell(4) != null) ? getCellValueAsString(newCatSheet.getRow(rowIndex).getCell(4)) : "";
        String catType = (newCatSheet.getRow(rowIndex) != null && newCatSheet.getRow(rowIndex).getCell(5) != null) ? getCellValueAsString(newCatSheet.getRow(rowIndex).getCell(5)) : "";
        String readyToAdopt = (newCatSheet.getRow(rowIndex) != null && newCatSheet.getRow(rowIndex).getCell(6) != null) ? getCellValueAsString(newCatSheet.getRow(rowIndex).getCell(6)) : "";
        String webDisplay = (newCatSheet.getRow(rowIndex) != null && newCatSheet.getRow(rowIndex).getCell(7) != null) ? getCellValueAsString(newCatSheet.getRow(rowIndex).getCell(7)) : "";
        String catStatus = (newCatSheet.getRow(rowIndex) != null && newCatSheet.getRow(rowIndex).getCell(8) != null) ? getCellValueAsString(newCatSheet.getRow(rowIndex).getCell(8)) : "";
        LocalDate statusChangeDate = (newCatSheet.getRow(rowIndex) != null && newCatSheet.getRow(rowIndex).getCell(9) != null) ? DateHelper.getCellValueAsDate(newCatSheet.getRow(rowIndex).getCell(9)) : null;
        String gender = (newCatSheet.getRow(rowIndex) != null && newCatSheet.getRow(rowIndex).getCell(10) != null) ? getCellValueAsString(newCatSheet.getRow(rowIndex).getCell(10)) : "";
        String catColor = (newCatSheet.getRow(rowIndex) != null && newCatSheet.getRow(rowIndex).getCell(11) != null) ? getCellValueAsString(newCatSheet.getRow(rowIndex).getCell(11)) : "";
        String chipType = (newCatSheet.getRow(rowIndex) != null && newCatSheet.getRow(rowIndex).getCell(12) != null) ? getCellValueAsString(newCatSheet.getRow(rowIndex).getCell(12)) : "";
        String chipNumber = (newCatSheet.getRow(rowIndex) != null && newCatSheet.getRow(rowIndex).getCell(13) != null) ? getCellValueAsString(newCatSheet.getRow(rowIndex).getCell(13)) : "";
        String tagNumber = (newCatSheet.getRow(rowIndex) != null && newCatSheet.getRow(rowIndex).getCell(14) != null) ? getCellValueAsString(newCatSheet.getRow(rowIndex).getCell(14)) : "";
        String DOBest = (newCatSheet.getRow(rowIndex) != null && newCatSheet.getRow(rowIndex).getCell(15) != null) ? getCellValueAsString(newCatSheet.getRow(rowIndex).getCell(15)) : "";
        LocalDate DOB = (newCatSheet.getRow(rowIndex) != null && newCatSheet.getRow(rowIndex).getCell(16) != null) ? DateHelper.getCellValueAsDate(newCatSheet.getRow(rowIndex).getCell(16)) :null;
        LocalDate altered = (newCatSheet.getRow(rowIndex) != null && newCatSheet.getRow(rowIndex).getCell(17) != null) ? DateHelper.getCellValueAsDate(newCatSheet.getRow(rowIndex).getCell(17)) : null;
        String genComments = (newCatSheet.getRow(rowIndex) != null && newCatSheet.getRow(rowIndex).getCell(18) != null) ? getCellValueAsString(newCatSheet.getRow(rowIndex).getCell(18)) : "";
        String hidCommments = (newCatSheet.getRow(rowIndex) != null && newCatSheet.getRow(rowIndex).getCell(19) != null) ? getCellValueAsString(newCatSheet.getRow(rowIndex).getCell(19)) : "";
        String distinctFeatures = (newCatSheet.getRow(rowIndex) != null && newCatSheet.getRow(rowIndex).getCell(20) != null) ? getCellValueAsString(newCatSheet.getRow(rowIndex).getCell(20)) : "";
        String shortBio = (newCatSheet.getRow(rowIndex) != null && newCatSheet.getRow(rowIndex).getCell(21) != null) ? getCellValueAsString(newCatSheet.getRow(rowIndex).getCell(21)) : "";
        String longBio = (newCatSheet.getRow(rowIndex) != null && newCatSheet.getRow(rowIndex).getCell(22) != null) ? getCellValueAsString(newCatSheet.getRow(rowIndex).getCell(22)) : "";
        String fearfulStat = (newCatSheet.getRow(rowIndex) != null && newCatSheet.getRow(rowIndex).getCell(23) != null) ? getCellValueAsString(newCatSheet.getRow(rowIndex).getCell(23)) : "";
        String childrenStat = (newCatSheet.getRow(rowIndex) != null && newCatSheet.getRow(rowIndex).getCell(24) != null) ? getCellValueAsString(newCatSheet.getRow(rowIndex).getCell(24)) : "";
        String animalStat = (newCatSheet.getRow(rowIndex) != null && newCatSheet.getRow(rowIndex).getCell(25) != null) ? getCellValueAsString(newCatSheet.getRow(rowIndex).getCell(25)) : "";
        String socialStat = (newCatSheet.getRow(rowIndex) != null && newCatSheet.getRow(rowIndex).getCell(26) != null) ? getCellValueAsString(newCatSheet.getRow(rowIndex).getCell(26)) : "";
        String declawedStat = (newCatSheet.getRow(rowIndex) != null && newCatSheet.getRow(rowIndex).getCell(27) != null) ? getCellValueAsString(newCatSheet.getRow(rowIndex).getCell(27)) : "";
        String litterTrainStat = (newCatSheet.getRow(rowIndex) != null && newCatSheet.getRow(rowIndex).getCell(28) != null) ? getCellValueAsString(newCatSheet.getRow(rowIndex).getCell(28)) : "";
        String specialNeedsStat = (newCatSheet.getRow(rowIndex) != null && newCatSheet.getRow(rowIndex).getCell(29) != null) ? getCellValueAsString(newCatSheet.getRow(rowIndex).getCell(29)) : "";

        if (hasData(newCatSheet, rowIndex)) {
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
            
            By submitButton = By.cssSelector("div[id='content'] input[id='submit']");
            String formUrl = driver.getCurrentUrl();
            driver.findElement(submitButton).click();
            waitForFormSubmission(driver, submitButton, formUrl, "New Intake - Add", rowIndex + 1);
            By similarAnimalsHeading = By.xpath(
                    "//h1[contains(normalize-space(.), 'Similar Animals')]"
                            + " | //h2[contains(normalize-space(.), 'Similar Animals')]");
            boolean similarAnimalsPage = driver.getTitle().contains("Similar Animals")
                    || driver.findElements(similarAnimalsHeading).stream().anyMatch(WebElement::isDisplayed);
            if (similarAnimalsPage) {
                throw new IllegalStateException(
                        "New Intake - Add row " + (rowIndex + 1)
                                + " needs review on the Similar Animals page; the spreadsheet row was retained.");
            }
        }
    }
}
