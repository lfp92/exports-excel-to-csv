package exceltocsv.strategy;

import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;

public interface SheetType {
    Workbook getWorkbook(String filePath) throws IOException;
}
