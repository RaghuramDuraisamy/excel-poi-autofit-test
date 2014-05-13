
package excel;

import java.awt.Font;
import java.awt.GraphicsEnvironment;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWriter {

    public static void main(String args[]) {
        String outputDir = "/tmp/";
        if (args.length > 0) {
            outputDir = args[0];
        }
        ExcelWriter ew = new ExcelWriter();
        try {
            ew.writeExcel(outputDir);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void writeExcel(String outputDir) throws IOException {
        printFonts();
        String templateFile = "/template.xlsx";
        XSSFWorkbook workbook = new XSSFWorkbook(this.getClass().getResourceAsStream(templateFile));
        for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
            XSSFSheet sheet = workbook.getSheetAt(sheetIndex);
            for (int columnIndex = 0; columnIndex < 10; columnIndex++) {
                sheet.autoSizeColumn(columnIndex);
            }
        }
        FileOutputStream out = new FileOutputStream(outputDir + "/poi_excel_" + System.currentTimeMillis() + ".xlsx");
        workbook.write(out);
    }

    private void printFonts() {
        System.out.println("\n\n\tAvailable Fonts:\n");
        Font[] fonts = GraphicsEnvironment.getLocalGraphicsEnvironment().getAllFonts();
        for (int i = 0; i < fonts.length; ++i) {
            System.out.print(fonts[i].getFontName() + " : ");
            System.out.print(fonts[i].getFamily() + " : ");
            System.out.println(fonts[i].getName());
        }
        System.out.println("\n\n\tAvailable Font-families:\n");
        String[] names = GraphicsEnvironment.getLocalGraphicsEnvironment().getAvailableFontFamilyNames();
        for (int i = 0; i < names.length; ++i) {
            System.out.println(names[i]);
        }
    }
}
