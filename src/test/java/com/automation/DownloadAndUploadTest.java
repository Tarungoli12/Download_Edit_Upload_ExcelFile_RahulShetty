package com.automation;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.io.*;
import java.time.Duration;
import java.util.Iterator;

public class DownloadAndUploadTest {

    String file = "C://Users//280713//Downloads//download.xlsx";
    String fruitName = "Mango"; //give exact name
    String updatedValue = "555";
    String columnName = "price";

    @Test
    public void test() throws IOException, InvalidFormatException {
        WebDriver driver = new ChromeDriver();
        driver.manage().window().maximize();
        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(5));
        driver.get("https://rahulshettyacademy.com/upload-download-test/");
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(3));

        //click on download
        driver.findElement(By.cssSelector("#downloadButton")).click();


        //edit the cell value in Excel file
        int rowIndex = getRowIndex(fruitName);
        int columnIndex = getColumnIndex(columnName);
        upDateCell(rowIndex, columnIndex, updatedValue);


        //upload the edited Excel file
        WebElement upload = driver.findElement(By.id("fileinput"));
        upload.sendKeys("C://Users//280713//Downloads//download.xlsx");

        //wait and verify the popup message and wait for popup message invisibility
        By uploadedMessage = By.cssSelector("div[class*='Toastify__toast-icon'] + div");
        wait.until(ExpectedConditions.visibilityOfElementLocated(uploadedMessage));
        Assert.assertEquals(driver.findElement(uploadedMessage).getText(), "Updated Excel Data Successfully.");
        wait.until(ExpectedConditions.invisibilityOfElementLocated(uploadedMessage));


        String indexOfPriceColumn = driver.findElement(By.xpath("//div[text()='Price']")).getAttribute("data-column-id");
        String fruitPrice = driver.findElement(By.xpath("//div[text()='" + fruitName + "']/ancestor::div[@role='row']//div[@data-column-id='" + indexOfPriceColumn + "']/div")).getText();
        System.out.println(fruitPrice);
        Assert.assertEquals(fruitPrice, updatedValue);
    }

    public int getRowIndex(String name) throws IOException{
        XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(file));
        XSSFSheet sheet = workbook.getSheetAt(0);
        XSSFRow firstRow = sheet.getRow(0);
        Iterator<Cell> cells = firstRow.cellIterator();
        int columnIndex = 0;
        while (cells.hasNext()) {
            Cell cell = cells.next();
            if (cell.getStringCellValue().equalsIgnoreCase("fruit_name")) {
                columnIndex = cell.getColumnIndex();
                break;
            }
        }
        int rowIndex = 0;
        Iterator<Row> rows = sheet.iterator();
        while (rows.hasNext()) {
            Row row = rows.next();
            if (row.getCell(columnIndex).getStringCellValue().equalsIgnoreCase(name)) {
                rowIndex = row.getRowNum();
            }
        }
        return rowIndex;
    }

    public int getColumnIndex(String name) throws IOException, InvalidFormatException {
        XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(file));
        XSSFSheet sheet = workbook.getSheet("Sheet1");
        XSSFRow row = sheet.getRow(0);
        Iterator<Cell> cells = row.cellIterator();
        int columnIndex = 0;
        while (cells.hasNext()) {
            Cell cell = cells.next();
            if (cell.getStringCellValue().equalsIgnoreCase(name)) {
                columnIndex = cell.getColumnIndex();
                break;
            }
        }
        return columnIndex;
    }

    public void upDateCell(int rowIndex, int columnIndex, String updatedValue) throws IOException, InvalidFormatException {
        XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(file));
        XSSFSheet sheet = workbook.getSheetAt(0);
        XSSFRow row = sheet.getRow(rowIndex);
        XSSFCell cell = row.getCell(columnIndex);
        cell.setCellValue(updatedValue);
        FileOutputStream stream = new FileOutputStream(file);
        workbook.write(stream);
        workbook.close();
        stream.close();
    }
}
