package Utils;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtils {
    public static XSSFWorkbook workbook;
    public static XSSFSheet sheet;

    public ExcelUtils(String excelPath, String sheetName) {
        try{
            workbook = new XSSFWorkbook(excelPath);
            sheet = workbook.getSheet(sheetName);
        }catch (Exception e){
            e.printStackTrace();
        }
    }

    public int getRowCount() {
        int rowCount = 0;
        try {
            rowCount = sheet.getPhysicalNumberOfRows();
            System.out.println("Number of rows: " + rowCount);
        } catch (Exception e) {
            System.out.println(e.getMessage());
            System.out.println(e.getCause());
            e.printStackTrace();
        }
        return rowCount;
    }

    public int getColumnCount() {
        int columnCount = 0;
        try {
            columnCount = sheet.getRow(0).getPhysicalNumberOfCells();
            System.out.println("Number of columns: " + columnCount);
        } catch (Exception e) {
            System.out.println(e.getMessage());
            System.out.println(e.getCause());
            e.printStackTrace();
        }
        return columnCount;
    }

    public String getCellDataString(int rowNum, int colNum) {
        String cellData = null;
        try{
            cellData = sheet.getRow(rowNum).getCell(colNum).getStringCellValue();
            System.out.println("This is a cell data: " + cellData);
        }catch (Exception e){
            System.out.println(e.getMessage());
            System.out.println(e.getCause());
            e.printStackTrace();
        }
        return cellData;
    }

    public void getCellDataNumber(int rowNum, int colNum) {
        try{
            double cellData = sheet.getRow(rowNum).getCell(colNum).getNumericCellValue();
            System.out.println("This is a cell data: " + cellData);
        }catch (Exception e){
            System.out.println(e.getMessage());
            System.out.println(e.getCause());
            e.printStackTrace();
        }
    }
}
