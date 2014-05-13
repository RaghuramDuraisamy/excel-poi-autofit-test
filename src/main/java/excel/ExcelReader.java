
package excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {

    public static void main(String args[]) {
        if (args.length < 1) {
            System.out.println("Provide input file path");
            System.exit(1);
        }
        String inputFile = args[0];
        ExcelReader ew = new ExcelReader();
        try {
            ew.readExcel(inputFile);
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    private void readExcel(String inputFile) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(new File(inputFile)));
        for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
            XSSFSheet sheet = workbook.getSheetAt(sheetIndex);
            for (int columnIndex = 0; columnIndex < 8; columnIndex++) {
                System.out.println(sheet.getColumnWidth(columnIndex));
            }
        }
    }

}
