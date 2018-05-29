/**
 *
 */
import ExcelReadWrite.*;
import java.io.IOException;

public class Runners {
    public static void main(String args[]) throws IOException {
        ExcelReader.readXLSXFile("files/excelfile.xlsx");
    }
}
