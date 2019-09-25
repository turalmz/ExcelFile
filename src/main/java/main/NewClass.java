/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package main;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import static org.apache.poi.hssf.usermodel.HeaderFooter.file;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author muzaffarlit
 */
public class NewClass {
    public static void main(String file) throws IOException, InvalidFormatException{
       // String file = "qqq.xlsx";
    
        //try {
           // POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(file));
            OPCPackage pkg = OPCPackage.open(new File(file));
            //HSSFWorkbook wb = new HSSFWorkbook(pkg);
            XSSFWorkbook wb = new XSSFWorkbook(pkg);
            XSSFSheet sheet = wb.getSheetAt(0);
            XSSFRow row;
            //XSSFCell cell;

            int rows; // No of rows
            rows = sheet.getPhysicalNumberOfRows();

            int cols = 0; // No of columns
            int tmp = 0;

            // This trick ensures that we get the data properly even if it doesn't start from first few rows
            for(int i = 0; i < 10 || i < rows; i++) {
                row = sheet.getRow(i);
                if(row != null) {
                    tmp = sheet.getRow(i).getPhysicalNumberOfCells();
                    if(tmp > cols) cols = tmp;
                }
            }

            for(int r = 0; r < rows; r++) {
                row = sheet.getRow(r);
                parseRow(row,cols);
            }
       // } catch(Exception ioe) {

       // }
    }
    
    
   static void parseRow(XSSFRow row, int cols){
       XSSFCell cell;
        if(row != null) {
        for(int c = 0; c < cols; c++) {
            cell = row.getCell((short)c);
            if(cell != null) {
                // Your code here

                System.out.print(cell+ " ");
                switch (cell.getCellType()) {
                case STRING: 
                    
                    break;
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {

                    } else {

                    }
                    break;
                case BOOLEAN:

                    break;
                case FORMULA:

                    break;
                case BLANK:

                default:
                    System.out.println("Could not determine cell type");
                    System.out.print(cell.getCellType());
                }

            }else{
                System.out.print(" ");
            }

        }
        System.out.println();
    }
    
    
    }
}
