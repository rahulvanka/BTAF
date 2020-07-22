package ExcelHandlers;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.IOException;

public class WorkbookHandler {

    public static Workbook getWorkbookObject(String pathToExcelFile) throws IOException, InvalidFormatException {
        return WorkbookFactory.create(new File(pathToExcelFile));
    }

    public static int getNumberOfSheetsInWorkbook(Workbook workbook){
        return workbook.getNumberOfSheets();
    }

}
