package tests;

import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.devtools.idealized.Javascript;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.security.Key;
import java.time.Duration;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.text.SimpleDateFormat;

public class FurstTest {
    public static void main(String[] args) throws IOException{
        WebDriver driver = new ChromeDriver();
        
        File DCinstructions = new File("src/main/resources/DCWorkbook.xlsm");
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
        //Create check-in record; will always be applicable for new cats
        if (checkinSheet != null) {
            createCheckInRecord(driver, checkinSheet);
        }
        //Add vaccinations if applicable
        if (vaccinations != null) {
            addVaccinations(driver, vaccSheet);
        }

        //Add medical check-up/weights if applicable
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
        String animalStat = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(25) != null) ? getCellValueAsString(newCatSheet.getRow(1).getCell(26)) : "";
        String socialStat = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(26) != null) ? getCellValueAsString(newCatSheet.getRow(1).getCell(25)) : "";
        String declawedStat = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(27) != null) ? getCellValueAsString(newCatSheet.getRow(1).getCell(25)) : "";
        String litterTrainStat = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(28) != null) ? getCellValueAsString(newCatSheet.getRow(1).getCell(25)) : "";
        String specialNeedsStat = (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(29) != null) ? getCellValueAsString(newCatSheet.getRow(1).getCell(25)) : "";

        while (newCatSheet.getRow(1) != null && newCatSheet.getRow(1).getCell(0) != null) {
            driver.findElement(By.linkText("Add Animal")).click();
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
                //dobField.click();
                //dobField.clear();
                dobField.sendKeys(DOB.toString());
                dobField.sendKeys(Keys.TAB);
                closeDatePicker(driver, dobField);
                //driver.findElement(By.name("g")).click();
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
            deleteRow(newCatSheet, 1);
        }
    }

    private static void createCheckInRecord(WebDriver driver, Sheet checkinSheet) {
        // Implementation for creating a check-in record based on the data in checkinSheet
        String numberCats = (checkinSheet.getRow(1) != null && checkinSheet.getRow(1).getCell(0) != null) ? getCellValueAsString(checkinSheet.getRow(1).getCell(0)) : "";
        int numCats = Integer.parseInt(numberCats);
        String catNames = (checkinSheet.getRow(2) != null && checkinSheet.getRow(2).getCell(0) != null) ? getCellValueAsString(checkinSheet.getRow(2).getCell(0)) : "";
        List<String> catList = Arrays.asList(catNames.split("\\s*;\\s*"));
        String broughtBy = (checkinSheet.getRow(3) != null && checkinSheet.getRow(3).getCell(0) != null) ? getCellValueAsString(checkinSheet.getRow(3).getCell(0)) : "";
        String prevOwner = (checkinSheet.getRow(4) != null && checkinSheet.getRow(4).getCell(0) != null) ? getCellValueAsString(checkinSheet.getRow(4).getCell(0)) : "";
        String prevAgency = (checkinSheet.getRow(5) != null && checkinSheet.getRow(5).getCell(0) != null) ? getCellValueAsString(checkinSheet.getRow(5).getCell(0)) : "";
        String surrReason = (checkinSheet.getRow(6) != null && checkinSheet.getRow(6).getCell(0) != null) ? getCellValueAsString(checkinSheet.getRow(6).getCell(0)) : "";
        String checkinType = (checkinSheet.getRow(7) != null && checkinSheet.getRow(7).getCell(0) != null) ? getCellValueAsString(checkinSheet.getRow(7).getCell(0)) : "";
        LocalDate checkinDate = (checkinSheet.getRow(8) != null && checkinSheet.getRow(8).getCell(0) != null) ? getCellValueAsDate(checkinSheet.getRow(8).getCell(0)) :null;
        String checker = (checkinSheet.getRow(9) != null && checkinSheet.getRow(9).getCell(0) != null) ? getCellValueAsString(checkinSheet.getRow(9).getCell(0)) : "";
        String sheltAnCounts = (checkinSheet.getRow(10) != null && checkinSheet.getRow(10).getCell(0) != null) ? getCellValueAsString(checkinSheet.getRow(10).getCell(0)) : "";
        String genComments = (checkinSheet.getRow(11) != null && checkinSheet.getRow(11).getCell(0) != null) ? getCellValueAsString(checkinSheet.getRow(11).getCell(0)) : "";
        String hidComments = (checkinSheet.getRow(12) != null && checkinSheet.getRow(12).getCell(0) != null) ? getCellValueAsString(checkinSheet.getRow(12).getCell(0)) : "";
        String move = (checkinSheet.getRow(13) != null && checkinSheet.getRow(13).getCell(0) != null) ? getCellValueAsString(checkinSheet.getRow(13).getCell(0)) : "";
        String currentlyWith = (checkinSheet.getRow(14) != null && checkinSheet.getRow(14).getCell(0) != null) ? getCellValueAsString(checkinSheet.getRow(14).getCell(0)) : "";
        String location = (checkinSheet.getRow(15) != null && checkinSheet.getRow(15).getCell(0) != null) ? getCellValueAsString(checkinSheet.getRow(15).getCell(0)) : "";
        String subLocation = (checkinSheet.getRow(16) != null && checkinSheet.getRow(16).getCell(0) != null) ? getCellValueAsString(checkinSheet.getRow(16).getCell(0)) : "";
        LocalDate today = LocalDate.now();

        while (checkinSheet.getRow(1) != null && checkinSheet.getRow(1).getCell(0) != null) {
            if (catList.size() != numCats) {
                System.out.println("Warning: Number of cats does not match the list of cat names.");
            } else {
                driver.findElement(By.id("s")).sendKeys(catList.get(0));
                driver.findElement(By.id("submit")).click();
                driver.findElement(By.xpath("//a[@class='animal'][text()='" + catList.get(0) + "']")).click();
                driver.findElement(By.cssSelector("a[title='Check-In History']")).click();
                driver.findElement(By.cssSelector("a[title='Add a Check-In']")).click();
                
                //Add a cat entry for each cat in the list and then fill its name
                for (int i = 1; i <= numCats-1; i++) {
                    driver.findElement(By.cssSelector("a[title='Add an additional animal to be checked in']")).click();
                    driver.findElement(By.name("aa" + i)).sendKeys(catList.get(i));
                }
            }

            driver.findElement(By.name("bb")).sendKeys(broughtBy);
            driver.findElement(By.name("po")).sendKeys(prevOwner);
            driver.findElement(By.name("j")).sendKeys(prevAgency);
            driver.findElement(By.name("sr")).sendKeys(surrReason);
            driver.findElement(By.name("ct")).sendKeys(checkinType);
            if (checkinDate != null) {
                driver.findElement(By.name("cd")).click();
                driver.findElement(By.name("cd")).sendKeys(Keys.chord(Keys.CONTROL + "a"), Keys.DELETE);
                driver.findElement(By.name("cd")).sendKeys(checkinDate.toString());
                driver.findElement(By.name("cd")).sendKeys(Keys.TAB);
            } else {
                driver.findElement(By.name("cd")).click();
                driver.findElement(By.name("cd")).sendKeys(Keys.chord(Keys.CONTROL + "a"), Keys.DELETE);
                driver.findElement(By.name("cd")).sendKeys(today.toString());
                driver.findElement(By.name("cd")).sendKeys(Keys.TAB);
            }
            driver.findElement(By.name("cb")).sendKeys(checker);
            driver.findElement(By.name("sacs")).sendKeys(sheltAnCounts);
            driver.findElement(By.name("gc")).sendKeys(genComments);
            driver.findElement(By.name("hc")).sendKeys(hidComments);
            
            //Process a move
            if ("Yes".equals(move)) {
                driver.findElement(By.name("mrc")).click();
                driver.findElement(By.name("mrcw")).sendKeys(currentlyWith);
                driver.findElement(By.name("mrl")).sendKeys(location);
                driver.findElement(By.name("mrsl")).sendKeys(subLocation);
            }
            
            driver.findElement(By.cssSelector("input[value='Add']")).click();

            deleteRow(checkinSheet, 1);
        }

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

        while (vaccSheet.getRow(1) != null && vaccSheet.getRow(1).getCell(0) != null) {
            driver.findElement(By.id("s")).sendKeys(catName);
            driver.findElement(By.id("submit")).click();
            driver.findElement(By.xpath("//a[@class='animal'][text()='" + catName + "']")).click();
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

            deleteRow(vaccSheet, 1);
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

        while (medCheckSheet.getRow(1) != null && medCheckSheet.getRow(1).getCell(0) != null) {
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
        
            deleteRow(medCheckSheet, 1);
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

        while (procedureSheet.getRow(1) != null && procedureSheet.getRow(1).getCell(0) != null) {
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

            deleteRow(procedureSheet, 1);
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

        while (prescriptionSheet.getRow(1) != null && prescriptionSheet.getRow(1).getCell(0) != null) {
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

            deleteRow(prescriptionSheet, 1);
        }
    }

    private static void closeDatePicker(WebDriver driver, WebElement field) {
        ((JavascriptExecutor) driver).executeScript(
                "var el = arguments[0]; if (el) { try { el.blur(); } catch (e) {} }" +
                        "var selectors = ['.ui-datepicker', '.datepicker', '.flatpickr-calendar', '.xdsoft_datetimepicker'];" +
                        "for (var i = 0; i < selectors.length; i++) {" +
                        "  var nodes = document.querySelectorAll(selectors[i]);" +
                        "  for (var j = 0; j < nodes.length; j++) {" +
                        "    nodes[j].style.display = 'none';" +
                        "    nodes[j].classList.remove('open');" +
                        "    nodes[j].classList.remove('show');" +
                        "  }" +
                        "}" +
                        "var overlays = document.querySelectorAll('.ui-datepicker-overlay, .ui-widget-overlay');" +
                        "for (var k = 0; k < overlays.length; k++) { overlays[k].style.display = 'none'; }" +
                        "if (window.jQuery) { try { window.jQuery('.ui-datepicker, .datepicker, .flatpickr-calendar, .xdsoft_datetimepicker').hide(); } catch (e) {} }",
                field);
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
