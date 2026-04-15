package tests;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.time.Duration;
import java.text.SimpleDateFormat;

public class FurstTest {
    public static void main(String[] args) throws IOException{
        WebDriver driver = new ChromeDriver();
        
        File DCinstructions = new File("TCHproject/src/main/resources/DCWorkbook.xlsm");
        FileInputStream fis = new FileInputStream(DCinstructions);
        Workbook workbook = new XSSFWorkbook(fis);
        
        Sheet credSheet = workbook.getSheet("Instructions");
        Sheet newCatSheet = workbook.getSheet("New Intake - Add");
        Sheet checkinSheet = workbook.getSheet("New Intake - Check-In");
        Sheet vaccSheet = workbook.getSheet("Vaccinations");
        Sheet medCheckSheet = workbook.getSheet("Medical Check-Up");
        Sheet procedureSheet = workbook.getSheet("Medical Procedure");
        Sheet prescriptionSheet = workbook.getSheet("Prescriptions");

        driver.get("https://ishelter.demoshelters.com/as/");

        // Read credentials from credSheet
        String username = credSheet.getRow(1).getCell(1).getStringCellValue();
        String password = credSheet.getRow(2).getCell(1).getStringCellValue();

        // Find elements and enter credentials
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
        wait.until(ExpectedConditions.presenceOfElementLocated(By.id("email")));

        driver.findElement(By.id("email")).sendKeys(username);
        driver.findElement(By.id("password")).sendKeys(password);
        driver.findElement(By.id("submit")).click();
        //driver.findElement(By.id("/as/")).click();

        //Create variables based on each relevant cell to determine what path to follow
                // Check if row exists before accessing cell
        Cell newCat = newCatSheet.getRow(1) != null ? newCatSheet.getRow(1).getCell(0) : null;
        Cell vaccinations = vaccSheet.getRow(1) != null ? vaccSheet.getRow(1).getCell(0) : null;
        Cell medCheck = medCheckSheet.getRow(1) != null ? medCheckSheet.getRow(1).getCell(0) : null;
        Cell procedure = procedureSheet.getRow(1) != null ? procedureSheet.getRow(1).getCell(0) : null;
        Cell prescriptions = prescriptionSheet.getRow(1) != null ? prescriptionSheet.getRow(1).getCell(0) : null;
        
        //Add new cat, create check-in record, add litter if applicable
        if (newCat != null) {
            addNewCat(driver, newCatSheet);
        }
        
        //Add vaccinations if applicable
        if (vaccinations != null) {
            addVaccinations(driver, vaccSheet);
        }

        //Add medical check-up if applicable
        if (medCheck != null) {
            addMedCheck(driver, medCheckSheet);
        }

        //Add medical procedure if applicable
        if (procedure != null) {
            addProcedure(driver, procedureSheet);
        }

        //Add prescriptions if applicable
        if (prescriptions != null) {
            addPrescriptions(driver, prescriptionSheet);
        }

        workbook.close();
        //driver.quit();
    }

    private static void addNewCat(WebDriver driver, Sheet newCatSheet) {
        // Implementation for adding a new cat based on the data in newCatSheet
        
        //Gather all filled fields for a new cat
        String nameofCat = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(0) != null) ? newCatSheet.getRow(1).getCell(0).getStringCellValue() : "";
        String AliasofCat = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(1) != null) ? newCatSheet.getRow(1).getCell(1).getStringCellValue() : "";
        String species = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(2) != null) ? newCatSheet.getRow(1).getCell(2).getStringCellValue() : "";
        String breed = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(3) != null) ? newCatSheet.getRow(1).getCell(3).getStringCellValue() : "";
        String catCode = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(4) != null) ? newCatSheet.getRow(1).getCell(4).getStringCellValue() : "";
        String readyToAdopt = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(5) != null) ? newCatSheet.getRow(1).getCell(5).getStringCellValue() : "";
        String webDisplay = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(6) != null) ? newCatSheet.getRow(1).getCell(6).getStringCellValue() : "";
        String catStatus = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(7) != null) ? newCatSheet.getRow(1).getCell(7).getStringCellValue() : "";
        String statusChangeDate = "";
        if (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(8) != null) {
            Cell scdCell = newCatSheet.getRow(1).getCell(8);
            if (scdCell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(scdCell)) {
                statusChangeDate = new SimpleDateFormat("MM/dd/yyyy").format(scdCell.getDateCellValue());
            }
        }
        String gender = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(9) != null) ? newCatSheet.getRow(1).getCell(9).getStringCellValue() : "";
        String catColor = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(10) != null) ? newCatSheet.getRow(1).getCell(10).getStringCellValue() : "";
        String chipType = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(11) != null) ? newCatSheet.getRow(1).getCell(11).getStringCellValue() : "";
        String chipNumber = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(12) != null) ? newCatSheet.getRow(1).getCell(12).getStringCellValue() : "";
        String tagNumber = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(13) != null) ? newCatSheet.getRow(1).getCell(13).getStringCellValue() : "";
        String DOBest = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(14) != null) ? newCatSheet.getRow(1).getCell(14).getStringCellValue() : "";
        String DOB = "";
        if (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(15) != null) {
            Cell dobCell = newCatSheet.getRow(1).getCell(15);
            if (dobCell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(dobCell)) {
                DOB = new SimpleDateFormat("MM/dd/yyyy").format(dobCell.getDateCellValue());
            }
        }
        String altered = "";
        if (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(16) != null) {
            Cell alteredCell = newCatSheet.getRow(1).getCell(16);
            if (alteredCell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(alteredCell)) {
                altered = new SimpleDateFormat("MM/dd/yyyy").format(alteredCell.getDateCellValue());
            }
        }
        String genComments = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(17) != null) ? newCatSheet.getRow(1).getCell(17).getStringCellValue() : "";
        String hidCommments = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(18) != null) ? newCatSheet.getRow(1).getCell(18).getStringCellValue() : "";
        String distinctFeatures = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(19) != null) ? newCatSheet.getRow(1).getCell(19).getStringCellValue() : "";
        String shortBio = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(20) != null) ? newCatSheet.getRow(1).getCell(20).getStringCellValue() : "";
        String longBio = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(21) != null) ? newCatSheet.getRow(1).getCell(21).getStringCellValue() : "";
        String behavior = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(22) != null) ? newCatSheet.getRow(1).getCell(22).getStringCellValue() : "";
        String attributes = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(23) != null) ? newCatSheet.getRow(1).getCell(23).getStringCellValue() : "";
        

        driver.findElement(By.linkText("Add Animal")).click();
        driver.findElement(By.name("n")).sendKeys(nameofCat);
        driver.findElement(By.name("as")).sendKeys(AliasofCat);
        driver.findElement(By.cssSelector("div[id='content'] select[name='s']")).sendKeys(species);
        driver.findElement(By.name("b")).sendKeys(breed);
        driver.findElement(By.name("c")).sendKeys(catCode);
        driver.findElement(By.name("rta")).click();
        driver.findElement(By.name("sw")).click();
        driver.findElement(By.name("st")).sendKeys(catStatus);
        driver.findElement(By.name("std")).sendKeys(statusChangeDate);
        driver.findElement(By.name("g")).sendKeys(gender);
        driver.findElement(By.name("pc")).sendKeys(catColor);
        driver.findElement(By.name("mt")).sendKeys(chipType);
        driver.findElement(By.name("mn")).sendKeys(chipNumber);
        driver.findElement(By.name("tn")).sendKeys(tagNumber);
        driver.findElement(By.name("e")).click();
        driver.findElement(By.name("bd")).sendKeys(DOB);
        driver.findElement(By.name("nd")).sendKeys(altered);
        driver.findElement(By.name("gc")).sendKeys(genComments);
        driver.findElement(By.name("hc")).sendKeys(hidCommments);
        driver.findElement(By.name("df")).sendKeys(distinctFeatures);
        driver.findElement(By.name("sbio")).sendKeys(shortBio);
        driver.findElement(By.name("lbio")).sendKeys(longBio);
        //driver.findElement(By.name("bh")).sendKeys(behavior);
        //driver.findElement(By.name("at")).sendKeys(attributes);
        //WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
        //wait.until(ExpectedConditions.elementToBeClickable(By.id("submit")));
        
        driver.findElement(By.cssSelector("div[id='content'] input[id='submit']")).click();
    }

    private static void addVaccinations(WebDriver driver, Sheet vaccSheet) {
        // Implementation for adding vaccinations based on the data in vaccSheet
    }

    private static void addMedCheck(WebDriver driver, Sheet medCheckSheet) {
        // Implementation for adding a medical check-up based on the data in medCheckSheet
    }

    private static void addProcedure(WebDriver driver, Sheet procedureSheet) {
        // Implementation for adding a medical procedure based on the data in procedureSheet
    }

    private static void addPrescriptions(WebDriver driver, Sheet prescriptionSheet) {
        // Implementation for adding prescriptions based on the data in prescriptionSheet
    }

}
