
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelMain {
    public static void main(String args[]) throws IOException {
        ExcelMain excelmain = new ExcelMain();
        excelmain.createAndSaveExcel();
    }

    private void createAndSaveExcel() throws IOException {
        Workbook xlsxworkbook = new XSSFWorkbook(); // represents the excel
        Sheet sheet1 = xlsxworkbook.createSheet("Est");

        Row row1 = sheet1.createRow(0);
        row1.createCell(0).setCellValue("Header 1");
        row1.createCell(1).setCellValue("Header 2");
        row1.createCell(2).setCellValue("Header 3");
        Row row2 = sheet1.createRow(1);
        row2.createCell(0).setCellValue("Value 1");
        row2.createCell(1).setCellValue("Value 2");
        row2.createCell(2).setCellValue("Value 3");

        xlsxworkbook.write(new FileOutputStream("ExcelSample.xlsx"));


    }
}
