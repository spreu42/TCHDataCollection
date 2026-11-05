package helpers;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.text.SimpleDateFormat;
import java.time.Duration;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.YearMonth;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.Date;
import java.util.Locale;

public final class DateHelper {
    private DateHelper() {
    }

    public static LocalDate getCellValueAsDate(Cell dateCell) {
        if (!isDateCell(dateCell)) {
            return null;
        }
        Date date = dateCell.getDateCellValue();
        return date.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
    }

    public static boolean isDateCell(Cell cell) {
        return cell != null && cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell);
    }

    public static String formatCellDate(Cell dateCell) {
        return new SimpleDateFormat("MM/dd/yyyy").format(dateCell.getDateCellValue());
    }

    public static void enterDateText(WebDriver driver, String fieldName, LocalDate date, boolean tabAfter) {
        enterDateText(driver, fieldName, date, true, tabAfter);
    }

    public static void enterDateText(
            WebDriver driver, String fieldName, LocalDate date, boolean focusBeforeEntry, boolean tabAfter) {
        WebElement dateField = driver.findElement(By.name(fieldName));
        if (focusBeforeEntry) {
            dateField.click();
        }
        dateField.sendKeys(Keys.chord(Keys.CONTROL + "a"), Keys.DELETE);
        dateField.sendKeys(date.toString());
        if (tabAfter) {
            dateField.sendKeys(Keys.TAB);
        }
    }

    public static void selectDateFromPicker(WebDriver driver, WebElement dateField, LocalDate date) {
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
        By calendar = By.cssSelector("table.ui-datepicker-calendar");
        By title = By.cssSelector(".ui-datepicker-title");
        By monthSelect = By.cssSelector(".ui-datepicker-title select.ui-datepicker-month");
        By yearSelect = By.cssSelector(".ui-datepicker-title select.ui-datepicker-year");
        String targetInputId = dateField.getDomAttribute("id");
        if (targetInputId == null || targetInputId.isEmpty()) {
            throw new IllegalStateException("Date input is missing its id.");
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
        if (!matchesSelectedDate(date, enteredDate)) {
            throw new IllegalStateException(
                    "Could not select date. Expected " + date
                            + " but the field contains '" + enteredDate + "'.");
        }
    }

    private static boolean matchesSelectedDate(LocalDate expectedDate, String fieldValue) {
        if (fieldValue == null) {
            return false;
        }
        if (expectedDate.toString().equals(fieldValue)) {
            return true;
        }

        DateTimeFormatter dateTimeFormat = DateTimeFormatter.ofPattern("dd MMMM uuuu hh:mm a", Locale.ENGLISH);
        try {
            return expectedDate.equals(LocalDateTime.parse(fieldValue, dateTimeFormat).toLocalDate());
        } catch (DateTimeParseException e) {
            return false;
        }
    }

    private static String getDatePickerInputId(WebDriver driver) {
        Object inputId = ((JavascriptExecutor) driver).executeScript(
                "var cell = document.querySelector('.ui-datepicker-calendar td[onclick]');" +
                        "var handler = cell ? cell.getAttribute('onclick') : null;" +
                        "var match = handler ? handler.match(/_selectDay\\('#([^']+)'/) : null;" +
                        "return match ? match[1] : null;");
        return inputId == null ? null : inputId.toString();
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
}
