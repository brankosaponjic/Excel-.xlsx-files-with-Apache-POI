package readxlsxfile;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

public class ReadByApachePOI {

    private static final String PATH = "C:\\Users\\Branko\\Desktop\\Software\\documents\\addresses.xlsx";

    public static void main(String[] args) {

        try {
            FileInputStream file = new FileInputStream(new File(PATH));
            Workbook workbook = new XSSFWorkbook(file);
            DataFormatter dataFormatter = new DataFormatter();
            Iterator<Sheet> sheets = workbook.sheetIterator();
            while (sheets.hasNext()) {
                Sheet sheet = sheets.next();
                System.out.println("Sheet name is: " + sheet.getSheetName());
                System.out.println("-----------------------------------");
                Iterator<Row> rows = sheet.rowIterator();
                while (rows.hasNext()) {
                    Row row = rows.next();
                    Iterator<Cell> cells = row.cellIterator();
                    while (cells.hasNext()) {
                        Cell cell = cells.next();
                        String cellValue = dataFormatter.formatCellValue(cell);
                        System.out.print(cellValue + "\t");
                    }
                    System.out.println();
                }
            }
            workbook.close();

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }
}
