package exceltocsv.strategy.types;

import exceltocsv.strategy.SheetType;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.IOException;

public class Xls implements SheetType {
    @Override
    public Workbook getWorkbook(String filePath) throws IOException {
        return new HSSFWorkbook(new POIFSFileSystem(new File(filePath)));
    }
}
