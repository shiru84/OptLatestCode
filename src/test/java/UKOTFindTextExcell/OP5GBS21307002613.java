package UKOTFindTextExcell;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Assert;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class OP5GBS21307002613 {


    static XSSFWorkbook workbook;
    static XSSFSheet sheet;
    static String URL;
    static String TITLE;
    static WebDriver driver ;
    static String TEXT;
    static FileOutputStream fos;
    static String projectPath = System.getProperty("user.dir");
    static FileInputStream fis;
    static String XLS = "/exel/OP5GBS21307002613/data.xlsx";
    static String FILE = projectPath + XLS;
    static Row row;
    static WebElement okThanks;
    static WebElement removeButton;
    static JavascriptExecutor js = (JavascriptExecutor)driver;
    static String theSelector;

    static int TESTING_ROW = 0;
    static int TESTING_URL_CELL = 1;
    static int TESTING_DATA_CEL = 2;
    static int TEST_RESULT_CEL =  2;
    static int TEST_RESULT_CEL3 =  3;
    static int TEST_RESULT_CEL4 =  4;
    static int TEST_RESULT_CEL5 =  5;
    static int TEST_RESULT_CEL6 =  6;
    static int TEST_RESULT_CEL7 =  7;
    static int TEST_RESULT_CEL8 =  8;
    static int TEST_RESULT_CEL9 =  9;
    static int TEST_RESULT_CEL10 =  10;


    static int n;
    public static void main(String[] args) throws Exception {

        n = 1;

        for (int i = 0; i <18; i++) {
            activityTest();
        }
        System.out.println("all done now!");
        quit();

    }
    public static void activityTest() throws Exception {
        openBrowser();

        goToPage(TESTING_ROW + n, TESTING_URL_CELL);
        Thread.sleep(2000);
        findExpectedText(TESTING_ROW + n, TESTING_DATA_CEL );
        System.out.println(n);


        quit();
        n = n + 1;

    }

    public static void scrollClick () throws Exception {

        JavascriptExecutor js = (JavascriptExecutor) driver;
        Thread.sleep(2000);
        js.executeScript("document.body.style.zoom = '0.5'");
//        js.executeScript("window.scrollTo (0, " + drop + ") ");
    }
    public static void openBrowser() throws Exception {
        driver = new ChromeDriver();
        driver.manage().window().maximize();
    }
    public static void getTheSelector( int rowNum, int cellNum) throws Exception {
        try {
            workbook = new XSSFWorkbook(FILE);
            sheet = workbook.getSheetAt(0);
            theSelector = sheet.getRow(rowNum).getCell(cellNum).getStringCellValue();
        } catch (Exception exp) {
            System.out.println(exp.getMessage());
            System.out.println(exp.getCause());
            exp.printStackTrace();
        }

        System.out.println("theSelector found");}
    public static void goToPage( int rowNum, int cellNum) throws Exception {
        try {
            workbook = new XSSFWorkbook(FILE);
            sheet = workbook.getSheetAt(0);
            URL = sheet.getRow(rowNum).getCell(cellNum).getStringCellValue();
        } catch (Exception exp) {
            System.out.println(exp.getMessage());
            System.out.println(exp.getCause());
            exp.printStackTrace();
        }
        driver.get(URL);
        System.out.println("URL found - "+URL);}
    public static void getN(int rowNum, int cellNum) {
        try {
            sheet = workbook.getSheetAt(0);
            n = (int) sheet.getRow(rowNum).getCell(cellNum).getNumericCellValue();
        } catch (Exception exp) {
            System.out.println(exp.getMessage());
            System.out.println(exp.getCause());
            exp.printStackTrace();
        }
    }

    public static void writeFile(XSSFWorkbook workbook,String FILE ) throws IOException, InterruptedException {
        FileOutputStream fos = new FileOutputStream(FILE);
        workbook.write(fos);
        fos.close();
        workbook.close();
        System.out.println("File Writen");
    }


    public static void findExpectedText(int rowNum, int colNum) throws Exception{

        try{getTheSelector(TESTING_ROW + n, colNum);
            System.out.println(theSelector);
            String CartMessage = driver.findElement(By.cssSelector("div > cx-cart-item-list > div > div > cx-cart-item > div > div.cart-item__details > div")).getText();

            Assert.assertEquals(CartMessage, theSelector);
            fis = new FileInputStream(FILE);
            workbook = new XSSFWorkbook(fis);
            sheet = workbook.getSheetAt(0);
            row = sheet.getRow(rowNum);
            row.createCell(colNum+1).setCellValue("IN STOCK");

        } catch (Exception exp) {
            System.out.println(exp.getMessage());
            System.out.println(exp.getCause());
            exp.printStackTrace();
            fis = new FileInputStream(FILE);
            workbook = new XSSFWorkbook(fis);
            sheet = workbook.getSheetAt(0);
            row = sheet.getRow(rowNum);
            String ErrorMessage = driver.findElement(By.cssSelector("div.opened > div:nth-child(1) > span:nth-child(2)")).getText();
            row.createCell(colNum+1).setCellValue(ErrorMessage);
        }
        writeFile(workbook, FILE);
    }
    public static void quit(){
        driver.quit();
    }



}
