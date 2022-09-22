package Zadatak1;

import com.github.javafaker.Faker;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class Main {
    public static void main(String[] args) {
        ReadExel.Read("ime.xlsx");
        try{
            writeExel("domaci22.xlsx");
        }catch (FileNotFoundException ex) {
            System.out.println("FileNotFound");
        }catch (IOException ex)
        {
            System.out.println("File invalid");
        }
    }
    public static void writeExel(String fileName) throws FileNotFoundException, IOException {

            Faker faker=new Faker();
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet =workbook.createSheet("imena");

        for(int j=2;j<10;j++) {
            XSSFRow row = sheet.createRow(j);
            for (int i = 0; i < 2; i++) {
                XSSFCell cell = row.createCell(i);
                cell.setCellValue(faker.name().firstName());
                cell.setCellValue(faker.name().lastName());
                cell.setCellValue(cell.getStringCellValue() + " ");

            }
            System.out.println();
        }
        FileOutputStream fileOutputStream = new FileOutputStream(new File(fileName));
        workbook.write(fileOutputStream);
        fileOutputStream.close();


    }
}
