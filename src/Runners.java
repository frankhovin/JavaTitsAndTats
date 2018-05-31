/**
 *
 */
import ExcelReadWrite.*;
import java.io.IOException;

public class Runners {
    public static void main(String args[]) throws IOException {
        //ExcelReader.readWholeSheet("files/excelfile.xlsx", "Sheet1");

        String info[] = ExcelReader.getAdjacentCells(
                "files/excelfile.xlsx", "Sheet1", "Firstname Lastname");
        System.out.print(info[0] + "," + info[1]);


    }
}
