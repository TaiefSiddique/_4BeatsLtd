package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.format.TextStyle;
import java.util.List;
import java.util.Locale;

public class ExcelSearch {

    public static void main(String[] args) throws IOException {
        String filePath = "C://Games//Taief-Cyberpunk/Excel.xlsx";
        String driverPath = "C://chromedriver-win64/chromedriver.exe";
        System.setProperty("webdriver.chrome.driver", driverPath);

        String day = LocalDate.now().getDayOfWeek().getDisplayName(TextStyle.FULL, Locale.ENGLISH);

        FileInputStream fis = new FileInputStream(new File(filePath));
        Workbook wb = new XSSFWorkbook(fis);
        Sheet sh = wb.getSheet(day);

        if (sh == null) {
            System.out.println("No sheet for: " + day);
            fis.close();
            return;
        }

        WebDriver drv = new ChromeDriver();

        for (Row r : sh) {
            Cell kwCell = r.getCell(2);
            if (kwCell == null || kwCell.getCellType() != CellType.STRING)
                continue;

            String kw = kwCell.getStringCellValue();
            drv.get("https://www.google.com");
            WebElement box = drv.findElement(By.name("q"));
            box.sendKeys(kw);
            box.submit();

            try {
                Thread.sleep(2000);
            } catch (InterruptedException e) {
                e.printStackTrace();
            }

            List<WebElement> links = drv.findElements(By.tagName("a"));
            String longTxt = "";
            String shortTxt = null;

            for (WebElement l : links) {
                String txt = l.getText().trim();
                if (txt.isEmpty())
                    continue;

                if (txt.length() > longTxt.length())
                    longTxt = txt;
                if (shortTxt == null || txt.length() < shortTxt.length()) {
                    shortTxt = txt;
                }
            }

            Cell longCell = r.createCell(3);
            Cell shortCell = r.createCell(4);

            longCell.setCellValue(longTxt);
            shortCell.setCellValue(shortTxt);
        }

        drv.quit();
        fis.close();

        FileOutputStream fos = new FileOutputStream(new File(filePath));
        wb.write(fos);
        fos.close();
        wb.close();

        System.out.println("Done updating");
    }
}
