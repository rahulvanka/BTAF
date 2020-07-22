package ExcelHandlers;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.cellwalk.CellWalk;

import java.util.ArrayList;
import java.util.List;

public class SheetHandler {
    public static Sheet getSheetFromWorkBook(Workbook workbook, String sheetName){
        return workbook.getSheet(sheetName);
    }

    public static Sheet getSheetFromWorkBook(Workbook workbook, int sheetIndex){
        return workbook.getSheetAt(sheetIndex);
    }

    public static Cell[][] getDataInCellRangeFromSheet(Sheet sheet, String range){
        CellRangeAddress cellRange = CellRangeAddress.valueOf(range);
        CellWalk cellWalk = new CellWalk(sheet, cellRange);
        Cell[][] data = new Cell[cellRange.getLastRow()][cellRange.getLastColumn()];
        cellWalk.traverse(((cell, cellWalkContext) -> {
            data[cellWalkContext.getRowNumber()][cellWalkContext.getColumnNumber()] = cell;
        }));
        return data;
    }
}
