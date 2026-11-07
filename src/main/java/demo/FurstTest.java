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
import java.util.ArrayList;
import java.util.List;

import helpers.SpreadsheetRowHelper;


public class FurstTest {
    public static void main(String[] args) throws IOException {
        String username = WorkflowSupport.requireEnvironmentVariable("TCH_USER");
        String password = WorkflowSupport.requireEnvironmentVariable("TCH_PASSWORD");
        File workbookFile = new File("src/main/resources/DCWorkbook.xlsm");
        SpreadsheetRowHelper.ensureWorkbookClosed(workbookFile);
        Workbook loadedWorkbook;
        try (FileInputStream input = new FileInputStream(workbookFile)) {
            loadedWorkbook = new XSSFWorkbook(input);
        }
        try (Workbook workbook = loadedWorkbook) {
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

                List<String> addedAnimalNames = new ArrayList<>();
                SpreadsheetRowHelper.processRows(
                        newCatSheet,
                        workbook,
                        workbookFile,
                        rowIndex -> WorkflowSupport.hasData(newCatSheet, rowIndex),
                        rowIndex -> {
                            String animalName = WorkflowSupport
                                    .getCellValueAsString(WorkflowSupport.getCell(newCatSheet, rowIndex, 0))
                                    .trim();
                            NewAnimalWorkflow.addNewCat(driver, newCatSheet, rowIndex);
                            addedAnimalNames.add(animalName);
                        });

                String newlyAddedAnimalName = addedAnimalNames.size() == 1
                        ? addedAnimalNames.get(0)
                        : null;
                for (int rowIndex = 1; rowIndex <= checkinSheet.getLastRowNum(); rowIndex++) {
                    if (WorkflowSupport.hasData(checkinSheet, rowIndex)
                            && !WorkflowSupport.hasCheckInCatNames(checkinSheet, rowIndex)) {
                        System.err.println("Skipping check-in row " + (rowIndex + 1)
                                + ": the cat count is filled, but the cat-name list is blank.");
                    }
                }
                String[] directCheckInAnimalName = { newlyAddedAnimalName };
                SpreadsheetRowHelper.processRows(
                        checkinSheet,
                        workbook,
                        workbookFile,
                        rowIndex -> WorkflowSupport.hasCheckInCatNames(checkinSheet, rowIndex),
                        rowIndex -> {
                            CheckInWorkflow.createCheckInRecord(
                                    driver, checkinSheet, rowIndex, directCheckInAnimalName[0]);
                            directCheckInAnimalName[0] = null;
                        });
                SpreadsheetRowHelper.processRows(
                        vaccSheet,
                        workbook,
                        workbookFile,
                        rowIndex -> WorkflowSupport.hasMedicalRecordData(
                                vaccSheet, "Vaccinations", rowIndex),
                        rowIndex -> VaccinationWorkflow.addVaccinations(driver, vaccSheet, rowIndex));
                SpreadsheetRowHelper.processRows(
                        medCheckSheet,
                        workbook,
                        workbookFile,
                        rowIndex -> WorkflowSupport.hasMedicalRecordData(
                                medCheckSheet, "Medical Check-Up", rowIndex),
                        rowIndex -> MedicalCheckUpWorkflow.addMedCheck(driver, medCheckSheet, rowIndex));
                SpreadsheetRowHelper.processRows(
                        procedureSheet,
                        workbook,
                        workbookFile,
                        rowIndex -> WorkflowSupport.hasMedicalRecordData(
                                procedureSheet, "Medical Procedure", rowIndex),
                        rowIndex -> MedicalProcedureWorkflow.addProcedure(driver, procedureSheet, rowIndex));
                PrescriptionWorkflow.Context prescriptionContext = new PrescriptionWorkflow.Context();
                SpreadsheetRowHelper.processRows(
                        prescriptionSheet,
                        workbook,
                        workbookFile,
                        rowIndex -> WorkflowSupport.hasPrescriptionData(prescriptionSheet, rowIndex),
                        rowIndex -> prescriptionContext.addPrescription(driver, prescriptionSheet, rowIndex));
            } finally {
                driver.quit();
            }
        }
    }
}
