package helpers;

import org.openqa.selenium.By;
import org.openqa.selenium.StaleElementReferenceException;
import org.openqa.selenium.TimeoutException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.time.Duration;
import java.util.List;
import java.util.stream.Collectors;

public final class AutocompleteHelper {
    private AutocompleteHelper() {
    }

    public static void selectMatchingSuggestion(
            WebDriver driver, By suggestions, String expectedName, String context) {
        String normalizedName = expectedName == null ? "" : expectedName.trim();
        try {
            new WebDriverWait(driver, Duration.ofSeconds(10)).until(currentDriver -> {
                for (WebElement option : currentDriver.findElements(suggestions)) {
                    try {
                        if (option.isDisplayed() && matchesName(option.getText(), normalizedName)) {
                            option.click();
                            return true;
                        }
                    } catch (StaleElementReferenceException ignored) {
                        // The autocomplete list can be replaced while its suggestions are being read.
                    }
                }
                return null;
            });
        } catch (TimeoutException e) {
            List<String> availableSuggestions = driver.findElements(suggestions).stream()
                    .map(option -> {
                        try {
                            return option.isDisplayed() ? option.getText().trim() : null;
                        } catch (StaleElementReferenceException ignored) {
                            return null;
                        }
                    })
                    .filter(text -> text != null)
                    .collect(Collectors.toList());
            throw new IllegalStateException(
                    "Could not select '" + normalizedName + "' for " + context
                            + ". Available suggestions: " + availableSuggestions,
                    e);
        }
    }

    private static boolean matchesName(String suggestionText, String expectedName) {
        String normalizedSuggestion = suggestionText == null ? "" : suggestionText.trim();
        return normalizedSuggestion.equalsIgnoreCase(expectedName)
                || normalizedSuggestion.regionMatches(
                        true, 0, expectedName + " (", 0, expectedName.length() + 2);
    }
}
