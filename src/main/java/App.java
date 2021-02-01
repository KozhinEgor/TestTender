import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.functions.Hyperlink;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;

public class App {
    public static void main(String[] args) throws IOException, Exception {
        InputStream ExcelFileToRead = new FileInputStream("C:\\Users\\Егор\\Documents\\работа\\2020_wk40.xlsx");
        XSSFWorkbook workbookRead = new XSSFWorkbook(ExcelFileToRead);
        XSSFSheet sheet = workbookRead.getSheetAt(0);

        InputStream ExcelFileToWrite = new FileInputStream("C:\\Users\\Егор\\Documents\\работа\\номенклатура.xlsx");
        XSSFWorkbook workbookWrite = new XSSFWorkbook(ExcelFileToWrite);
        int n =1;
        workbookWrite.getSheetAt(0).getRow(0).;

        while(sheet.getRow(n).getCell(2) != null){
            String[] MasStr;
            MasStr = sheet.getRow(n).getCell(13).toString().split(";");
            for(String el:MasStr){
                System.out.println(el);
            }
            n++;
        }
        System.out.println(n);
        ExcelFileToRead.close();
    }

}
