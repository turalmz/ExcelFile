/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package main;

/**
 *
 * @author muzaffarlit
 */
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;  

public class YeniReportSplitter {

    private final String fileName;
    private final int maxRows;

    public YeniReportSplitter(String fileName, final int maxRows) {

        ZipSecureFile.setMinInflateRatio(0);

        this.fileName = fileName;
        this.maxRows = maxRows;

        try {
            /* Read in the original Excel file. */
            OPCPackage pkg = OPCPackage.open(new File(fileName));
            XSSFWorkbook workbook = new XSSFWorkbook(pkg);
            XSSFSheet sheet = workbook.getSheetAt(0);

            /* Only split if there are more rows than the desired amount. */
            if (sheet.getPhysicalNumberOfRows() >= maxRows) {
                List<SXSSFWorkbook> wbs = splitWorkbook(workbook);
                //writeWorkBooks(wbs);
            }
            pkg.close();
        }
        catch (EncryptedDocumentException | IOException | InvalidFormatException e) {
            e.printStackTrace();
        }
    }

    private List<SXSSFWorkbook> splitWorkbook(XSSFWorkbook workbook) {

        List<SXSSFWorkbook> workbooks = new ArrayList<SXSSFWorkbook>();

        SXSSFWorkbook wb = new SXSSFWorkbook();
        SXSSFSheet sh = (SXSSFSheet) wb.createSheet();

        SXSSFRow newRow;
        SXSSFCell newCell;

        int rowCount = 0;
        int colCount = 0;

        XSSFSheet sheet = workbook.getSheetAt(0);

        for (Row row : sheet) {
            newRow = (SXSSFRow) sh.createRow(rowCount++);

            /* Time to create a new workbook? */
            if (rowCount == maxRows) {
                workbooks.add(wb);
                wb = new SXSSFWorkbook();
                sh = (SXSSFSheet) wb.createSheet();
                rowCount = 0;
            }

            for (Cell cell : row) {
                newCell = (SXSSFCell) newRow.createCell(colCount++);
                newCell = setValue(newCell, cell);

                CellStyle newStyle = wb.createCellStyle();
                newStyle.cloneStyleFrom(cell.getCellStyle());
                newCell.setCellStyle(newStyle);
            }
            colCount = 0;
        }

        /* Only add the last workbook if it has content */
        if (wb.getSheetAt(0).getPhysicalNumberOfRows() > 0) {
            workbooks.add(wb);
        }
        return workbooks;
    }

    /*
     * Grabbing cell contents can be tricky. We first need to determine what
     * type of cell it is.
     */
    private SXSSFCell setValue(SXSSFCell newCell, Cell cell) {
        switch (cell.getCellType()) {
        case STRING: 
            newCell.setCellValue(cell.getRichStringCellValue().getString());
            break;
        case NUMERIC:
            if (DateUtil.isCellDateFormatted(cell)) {
                newCell.setCellValue(cell.getDateCellValue());
            } else {
                newCell.setCellValue(cell.getNumericCellValue());
            }
            break;
        case BOOLEAN:
            newCell.setCellValue(cell.getBooleanCellValue());
            break;
        case FORMULA:
            newCell.setCellFormula(cell.getCellFormula());
            break;
        case BLANK:
            newCell.setCellValue("");
        default:
            System.out.println("Could not determine cell type");
            System.out.print(cell.getCellType());
        }
        return newCell;
    }

    /* Write all the workbooks to disk. */


    public static void main(String[] args) throws IOException{
        /* This will create a new workbook every 1000 rows. */
        
       System.out.println("Working Directory = " +
              System.getProperty("user.dir"));
        
        new YeniReportSplitter("qqq.xlsx", 1000);
    }

}