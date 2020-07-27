package ExcelHandlers;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

public class ApacheExcel {

    public static void main(String[] args) throws IOException {

        String fileName = "C:\\Users\\RX854WN\\OneDrive - EY\\Documents\\Microsoft\\hssftest.xls";

        InputStream input = new FileInputStream(fileName);

        HSSFWorkbook wb = new HSSFWorkbook(input);
        HSSFSheet sheet = wb.getSheetAt(0);

        getRowNum(sheet, 1, "b2" );

        System.out.println(getRowNum(sheet, 1, "b2" ));

        System.out.println(getColNum(sheet, 1, "b2"));
    }

    private static HSSFCell getCellContent(HSSFSheet sheet, int rownr, int colnr) {

        HSSFRow row = sheet.getRow(rownr);
        HSSFCell cell = row.getCell(colnr);

        return cell;

    }

    private static int findRow(HSSFSheet sheet, String cellContent) {
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell.getCellType() == CellType.STRING) {
                    if (cell.getRichStringCellValue().getString().trim().equals(cellContent)) {
                        return row.getRowNum();
                    }
                }
            }
        }
        return 0;
    }

    private static int getRowNum(HSSFSheet sheet, int colnr, String cellContent) {

        int rownr;

        rownr = findRow(sheet, cellContent);

        HSSFCell cell = getCellContent(sheet, rownr, colnr);

        String cellCheck = cell.getStringCellValue();

        if (cellCheck.equals(cellContent)) {

            return rownr;
        }
        else {
            return -1;
        }

    }

    private static int getColNum(HSSFSheet sheet, int rownr, String cellContent) {

        Row row = sheet.getRow(rownr);

        for (Cell cell : row) {
            if (cell.getCellType() == CellType.STRING) {
                if (cell.getRichStringCellValue().getString().trim().equals(cellContent)) {
                    return cell.getColumnIndex();
                }


            }

        }
        return 0;
    }
}






