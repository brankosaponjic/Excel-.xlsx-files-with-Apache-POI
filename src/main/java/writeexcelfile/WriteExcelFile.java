package writeexcelfile;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;

public class WriteExcelFile {
    public static void main(String[] args) {
        try {
            // Create workbook in .xlsx format
            Workbook workbook = new XSSFWorkbook();
            // For .xls workbook use new HSSFWorkbook

            // Create Sheet
            Sheet sheet = workbook.createSheet("Expenses");

            // Create top row with column headings
            String[] columnHeadings = {"ID", "Type", "Amount", "Currency", "Date"};

            // Makeup for header :)
            Font headerFront = workbook.createFont();
            headerFront.setBold(true);
            headerFront.setFontHeightInPoints((short)12);
            headerFront.setColor(IndexedColors.BLACK.index);

            // Cell Style
            CellStyle headerStyle = workbook.createCellStyle();
            headerStyle.setFont(headerFront);
            headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            headerStyle.setAlignment(HorizontalAlignment.CENTER);
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);

            // Header row
            Row headerRow = sheet.createRow(0);
            // Iterate over the column heading to create columns
            for (int i=0; i<columnHeadings.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(columnHeadings[i]);
                cell.setCellStyle(headerStyle);
            }

            // Freeze header row
//            sheet.createFreezePane(0,1);
            // Fill data
            ArrayList<Expense> expenseArrayList = createData();
            CreationHelper creationHelper = workbook.getCreationHelper();
            CellStyle dateStyle = workbook.createCellStyle();
            dateStyle.setDataFormat(creationHelper.createDataFormat().getFormat("MM/dd/yyyy"));
            int rowNum = 1;
            for (Expense expense : expenseArrayList) {
                Row row = sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(expense.getId());
                row.createCell(1).setCellValue(expense.getType());
                row.createCell(2).setCellValue(expense.getAmount());
                row.createCell(3).setCellValue(expense.getCurrency());
                Cell dateCell = row.createCell(4);
                dateCell.setCellValue(expense.getDate());
                dateCell.setCellStyle(dateStyle);
            }

            // Group and collapse rows
            int numOfRows = sheet.getLastRowNum();
            sheet.groupRow(1, numOfRows);
            sheet.setRowGroupCollapsed(1,true);

            // Create Sum column
            Row sumRow = sheet.createRow(rowNum);
            Cell sumRowTitle = sumRow.createCell(0);
            sumRowTitle.setCellValue("Total");
            sumRowTitle.setCellStyle(headerStyle);

            String strFormula = "SUM(C2:C" + rowNum + ")";
            Cell sumCell = sumRow.createCell(2);
            sumCell.setCellFormula(strFormula);
            sumCell.setCellValue(true);

            Cell currencyCell = sumRow.createCell(3);
            currencyCell.setCellValue("RSD");

            // Autosize columns
            for (int i=0; i<columnHeadings.length; i++) {
                sheet.autoSizeColumn(i);
            }

            // Write workbook output to a file
            FileOutputStream fileOut = new FileOutputStream("C:\\Users\\Branko\\Desktop\\Software\\documents\\Expenses.xlsx");
            workbook.write(fileOut);
            workbook.close();
            System.out.println("Completed!");

        }catch(Exception e) {
            e.printStackTrace();
        }
    }

    private static ArrayList<Expense> createData() throws ParseException {
        ArrayList<Expense> arrayList = new ArrayList<>();
        arrayList.add(new Expense(1, "Coffee Shop", 1000, "RSD",
                      new SimpleDateFormat("MM/dd/yyyy").parse("09/01/2020")));
        arrayList.add(new Expense(2, "Supermarket", 8200, "RSD",
                      new SimpleDateFormat("MM/dd/yyyy").parse("09/02/2020")));
        arrayList.add(new Expense(3, "Coffee Shop", 300, "RSD",
                      new SimpleDateFormat("MM/dd/yyyy").parse("09/02/2020")));
        arrayList.add(new Expense(4, "Supermarket", 1290, "RSD",
                      new SimpleDateFormat("MM/dd/yyyy").parse("09/03/2020")));
        arrayList.add(new Expense(5, "Coffee Shop", 900, "RSD",
                      new SimpleDateFormat("MM/dd/yyyy").parse("09/03/2020")));
        arrayList.add(new Expense(6, "Gymnastics", 4500, "RSD",
                      new SimpleDateFormat("MM/dd/yyyy").parse("09/03/2020")));
        arrayList.add(new Expense(7, "School supplies", 240, "RSD",
                      new SimpleDateFormat("MM/dd/yyyy").parse("09/04/2020")));
        arrayList.add(new Expense(8, "Supermarket", 1450, "RSD",
                      new SimpleDateFormat("MM/dd/yyyy").parse("09/04/2020")));
        arrayList.add(new Expense(9, "Coffee Shop", 350, "RSD",
                      new SimpleDateFormat("MM/dd/yyyy").parse("09/04/2020")));
        arrayList.add(new Expense(10, "Grill", 800, "RSD",
                      new SimpleDateFormat("MM/dd/yyyy").parse("09/04/2020")));
        arrayList.add(new Expense(11, "Laptop", 119000, "RSD",
                      new SimpleDateFormat("MM/dd/yyyy").parse("09/05/2020")));
        arrayList.add(new Expense(12, "Snickers", 6000, "RSD",
                      new SimpleDateFormat("MM/dd/yyyy").parse("09/05/2020")));
        arrayList.add(new Expense(13, "Coffee Shop", 1100, "RSD",
                      new SimpleDateFormat("MM/dd/yyyy").parse("09/05/2020")));
        arrayList.add(new Expense(14, "Clothes", 5900, "RSD",
                      new SimpleDateFormat("MM/dd/yyyy").parse("09/05/2020")));
        arrayList.add(new Expense(15, "Footwear", 5660, "RSD",
                      new SimpleDateFormat("MM/dd/yyyy").parse("09/05/2020")));
        arrayList.add(new Expense(16, "Windows Installation", 2500, "RSD",
                      new SimpleDateFormat("MM/dd/yyyy").parse("09/05/2020")));
        arrayList.add(new Expense(17, "Coffee Shop", 1850, "RSD",
                      new SimpleDateFormat("MM/dd/yyyy").parse("09/06/2020")));
        arrayList.add(new Expense(18, "Supermarket", 3470, "RSD",
                      new SimpleDateFormat("MM/dd/yyyy").parse("09/06/2020")));
        arrayList.add(new Expense(19, "Supermarket", 790, "RSD",
                      new SimpleDateFormat("MM/dd/yyyy").parse("09/07/2020")));
        arrayList.add(new Expense(20, "Material", 23500, "RSD",
                      new SimpleDateFormat("MM/dd/yyyy").parse("09/07/2020")));
        arrayList.add(new Expense(21, "Breakfast", 200, "RSD",
                      new SimpleDateFormat("MM/dd/yyyy").parse("09/08/2020")));
        arrayList.add(new Expense(22, "Coffee Shop", 370, "RSD",
                      new SimpleDateFormat("MM/dd/yyyy").parse("09/08/2020")));
        arrayList.add(new Expense(22, "Supermarket", 4920, "RSD",
                      new SimpleDateFormat("MM/dd/yyyy").parse("09/08/2020")));
        arrayList.add(new Expense(23, "Clothes", 6500, "RSD",
                      new SimpleDateFormat("MM/dd/yyyy").parse("09/08/2020")));
        arrayList.add(new Expense(24, "Grill", 1960, "RSD",
                      new SimpleDateFormat("MM/dd/yyyy").parse("09/09/2020")));
        arrayList.add(new Expense(25, "Coffee Shop", 350, "RSD",
                      new SimpleDateFormat("MM/dd/yyyy").parse("09/09/2020")));
        arrayList.add(new Expense(26, "Pharmacy", 540, "RSD",
                      new SimpleDateFormat("MM/dd/yyyy").parse("09/09/2020")));
        arrayList.add(new Expense(27, "Charity Donation", 820, "RSD",
                      new SimpleDateFormat("MM/dd/yyyy").parse("09/09/2020")));
        return arrayList;

    }
}