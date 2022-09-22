package Zadatak1;

import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public  class ReadExel {

    public static void Read(String path) {


        {
            try {
               FileInputStream fileInputStream = new FileInputStream(new File(path));
                XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
                XSSFSheet sheet = workbook.getSheet("Sheet1");



                for(int j=0;j<2;j++)
                {
                    XSSFRow row= sheet.getRow(j);
                    for(int i=0;i<2;i++)
                    {
                        XSSFCell cell = row.getCell(i);
                        System.out.print(cell.getStringCellValue() + " ");
                    }
                    System.out.println();
                }

            } catch (FileNotFoundException e) {
                System.out.println("FileNotFound.class");

            } catch (IOException e) {
                e.printStackTrace();
            }catch (NullPointerException e)
            {
                //
            }
        }


    }


}
