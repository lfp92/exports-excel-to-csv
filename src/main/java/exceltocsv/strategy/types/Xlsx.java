package exceltocsv.strategy.types;

import exceltocsv.strategy.SheetType;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

public class Xlsx implements SheetType {
    @Override
    public Workbook getWorkbook(String filePath) throws IOException {
        return new XSSFWorkbook(new FileInputStream(filePath));
    }
}
