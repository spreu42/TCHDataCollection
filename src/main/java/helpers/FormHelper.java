package helpers;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.Select;

public final class FormHelper {
    private FormHelper() {
    }

    public static void enterText(WebDriver driver, String fieldName, String value) {
        driver.findElement(By.name(fieldName)).sendKeys(value);
    }

    public static void selectByVisibleText(WebDriver driver, String fieldName, String value) {
        new Select(driver.findElement(By.name(fieldName))).selectByVisibleText(value);
    }

    public static void selectByValue(WebDriver driver, String fieldName, String value) {
        new Select(driver.findElement(By.name(fieldName))).selectByValue(value);
    }

    public static void setCheckbox(WebDriver driver, String fieldName, boolean selected) {
        WebElement checkbox = driver.findElement(By.name(fieldName));
        if (checkbox.isSelected() != selected) {
            checkbox.click();
        }
    }
}
