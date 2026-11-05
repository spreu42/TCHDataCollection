package helpers;

import org.openqa.selenium.By;
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
            WebElement matchingSuggestion = new WebDriverWait(driver, Duration.ofSeconds(10))
                    .until(currentDriver -> currentDriver.findElements(suggestions).stream()
                            .filter(WebElement::isDisplayed)
                            .filter(option -> matchesName(option.getText(), normalizedName))
                            .findFirst()
                            .orElse(null));
            matchingSuggestion.click();
        } catch (TimeoutException e) {
            List<String> availableSuggestions = driver.findElements(suggestions).stream()
                    .filter(WebElement::isDisplayed)
                    .map(option -> option.getText().trim())
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
