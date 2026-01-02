package MA_MSG_Suite_OB;

import org.openqa.selenium.*;
import java.util.ArrayList;

public class SwitchTab {

    public static void tabswitch(WebDriver driver, int num) {

        if (driver == null) {
            throw new IllegalArgumentException("❌ WebDriver passed to tabswitch() is null");
        }

        // Get all open window handles
        ArrayList<String> tabs = new ArrayList<>(driver.getWindowHandles());

        if (tabs.size() <= num) {
            throw new IllegalArgumentException("❌ Requested tab index " + num + " but only " + tabs.size() + " tabs are open.");
        }

        // Switch to the desired tab
        driver.switchTo().window(tabs.get(num));

        // Optional: Bring tab to front
        JavascriptExecutor js = (JavascriptExecutor) driver;
        js.executeScript("window.focus();");

        System.out.println("✅ Switched to tab index " + num + " and attempted to bring it to the front.");
    }
}
