package demo;

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


public class FurstTest {
    public static void main(String[] args) throws IOException {
        String username = WorkflowSupport.requireEnvironmentVariable("TCH_USER");
        String password = WorkflowSupport.requireEnvironmentVariable("TCH_PASSWORD");
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
                if (WorkflowSupport.hasData(newCatSheet)) {
                    newlyAddedAnimalName = WorkflowSupport.getCellValueAsString(newCatSheet.getRow(1).getCell(0)).trim();
                    NewAnimalWorkflow.addNewCat(driver, newCatSheet);
                }
                boolean checkInCompleted = false;
                if (WorkflowSupport.hasCheckInCatNames(checkinSheet)) {
                    checkInCompleted = CheckInWorkflow.createCheckInRecord(driver, checkinSheet, newlyAddedAnimalName);
                } else if (WorkflowSupport.hasData(checkinSheet)) {
                    System.err.println(
                            "Skipping check-in: the cat count is filled, but the cat-name list is blank.");
                }
                if (WorkflowSupport.hasMedicalRecordData(vaccSheet, "Vaccinations")) {
                    VaccinationWorkflow.addVaccinations(driver, vaccSheet);
                }
                if (WorkflowSupport.hasMedicalRecordData(medCheckSheet, "Medical Check-Up")) {
                    MedicalCheckUpWorkflow.addMedCheck(driver, medCheckSheet);
                }
                if (WorkflowSupport.hasMedicalRecordData(procedureSheet, "Medical Procedure")) {
                    MedicalProcedureWorkflow.addProcedure(driver, procedureSheet);
                }
                if (WorkflowSupport.hasPrescriptionData(prescriptionSheet)) {
                    PrescriptionWorkflow.addPrescriptions(driver, prescriptionSheet);
                }
            } finally {
                driver.quit();
            }
        }
    }
}
