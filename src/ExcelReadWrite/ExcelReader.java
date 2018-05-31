package ExcelReadWrite;


import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

public class ExcelReader {
    /**
     * Reads the whole sheet (limited to?) and prints each cell value.
     *
     * @param filename
     * @param sheetname
     * @throws IOException
     */
    public static void readWholeSheet (String filename, String sheetname) throws IOException {
        InputStream ExcelFileToRead = new FileInputStream(filename);
        XSSFWorkbook wb = new XSSFWorkbook(ExcelFileToRead);

        XSSFSheet sheet = wb.getSheet(sheetname);
        XSSFRow row;
        XSSFCell cell;

        Iterator rows = sheet.rowIterator();

        while (rows.hasNext()) {
            row =(XSSFRow) rows.next();
            Iterator cells = row.cellIterator();
            while (cells.hasNext()) {
                cell=(XSSFCell) cells.next();

                System.out.print(cell.getStringCellValue()+" ");

                // If we need to know what data types are in the cells:
                // Deprecated:
                /*if (cell.getCellType() == XSSFCell.CELL_TYPE_STRING) {
                    System.out.print(cell.getStringCellValue()+" ");
                }
                else if(cell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC) {
                    System.out.print(cell.getNumericCellValue()+" ");
                } else {
                }*/
                // Use getCellTypeEnum() instead (below). The plan is to rename getCellTypeEnum()
                // back to getCellType() in POI 4.0.
                /*if (cell.getCellTypeEnum() == CellType.STRING) {
                    System.out.print(cell.getStringCellValue()+" ");
                } else if (cell.getCellTypeEnum() == CellType.NUMERIC)) {
                //    System.out.print(cell.getNumericCellValue()+" ");
                }*/

            }

            System.out.println();
        }
    }

    /**
     * Search the sheet for a specific string and get the values in the adjacent cells.
     *
     * @param filename
     * @param sheetname
     * @param str
     * @throws IOException
     */
    public static String[] getAdjacentCells (String filename, String sheetname, String str) throws IOException {
        InputStream ExcelFileToRead = new FileInputStream(filename);
        XSSFWorkbook wb = new XSSFWorkbook(ExcelFileToRead);

        XSSFSheet sheet = wb.getSheet(sheetname);
        XSSFRow row;
        XSSFCell cell;

        Iterator rows = sheet.rowIterator();
        String[] values = new String[2];

        while (rows.hasNext()) {
            row =(XSSFRow) rows.next();
            Iterator cells = row.cellIterator();
            while (cells.hasNext()) {
                cell = (XSSFCell) cells.next();
                if (cell.getStringCellValue().equals(str)) {
                    cell = (XSSFCell) cells.next();
                    values[0] = cell.getStringCellValue();
                    cell = (XSSFCell) cells.next();
                    values[1] = cell.getStringCellValue();

                    return values;
                }
            }

        }

        return new String[2];
    }









}
