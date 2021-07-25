import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class ExcelReader
{
    public static void main(String[] args) throws IOException {
        ExcelReader excelReader = new ExcelReader();
        excelReader.readexcel();
    }
    private void readexcel() throws IOException {
        Workbook excelworkbook = new XSSFWorkbook(new FileInputStream("ExcelSample.xlsx"));
        int numberofsheets = excelworkbook.getNumberOfSheets();
        for(int i=0;i<numberofsheets;i++)
        {
            Sheet sheet = excelworkbook.getSheetAt(i);
           int rowcount = sheet.getPhysicalNumberOfRows();
           for(int j=0;j<rowcount;j++)
           {
               Row row = sheet.getRow(j);
               int cellcount = row.getPhysicalNumberOfCells();
               for(int k=0;k<cellcount;k++)
               {
                   System.out.println(row.getCell(k).toString() + "\t\t");
               }
               System.out.println("END");
           }
        }

    }
}
