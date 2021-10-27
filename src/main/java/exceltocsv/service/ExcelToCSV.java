package exceltocsv.service;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.PrintStream;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ExcelToCSV {
    private static final Pattern RXQUOTE = Pattern.compile("\"");
    private String fileName;

    private static String encodeValue(String value) {
        boolean needQuotes = false;
        if (value.indexOf(',') != -1 || value.indexOf('"') != -1 || value.indexOf('\n') != -1
                || value.indexOf('\r') != -1)
            needQuotes = true;
        Matcher m = RXQUOTE.matcher(value);
        if (m.find())
            needQuotes = true;
        value = m.replaceAll("\"\"");
        if (needQuotes)
            return "\"" + value + "\"";
        else
            return value;
    }

    public void exportCSV(String inputFilePath, String outputFilePath) {
        this.fileName = inputFilePath;
        String arr[] = fileName.split("(\\.)");
        String ext = arr[arr.length - 1];
        Workbook wb = null;
        PrintStream out = null;
        DataFormatter formatter = new DataFormatter();
        try {

            byte[] bom = {(byte) 0xEF, (byte) 0xBB, (byte) 0xBF};
            FileInputStream file = new FileInputStream(new File(this.fileName));
            if (ext.equals("xlsx")) {
                wb = new XSSFWorkbook(file);
            } else {
                wb = new HSSFWorkbook(new POIFSFileSystem(new File(this.fileName)));
            }
            FormulaEvaluator fe = wb.getCreationHelper().createFormulaEvaluator();
            {
                for (int sheetNo = 0, ns = wb.getNumberOfSheets(); sheetNo < ns; sheetNo++) {
                    Sheet sheet = wb.getSheetAt(sheetNo);

                    out = new PrintStream(new FileOutputStream(new File(outputFilePath + "_" + sheet.getSheetName() + ".csv")), true, "UTF-8");
                    out.println(bom);
                    for (int r = 0, rn = sheet.getLastRowNum(); r <= rn; r++) {
                        Row row = sheet.getRow(r);
                        if (row == null) {
                            out.println(',');
                            continue;
                        }
                        boolean firstCell = true;
                        for (int c = 0, cn = row.getLastCellNum(); c < cn; c++) {
                            Cell cell = row.getCell(c, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                            if (!firstCell)
                                out.print(',');
                            if (cell != null) {
                                if (fe != null)
                                    cell = fe.evaluateInCell(cell);
                                String value = formatter.formatCellValue(cell);
                                if (cell.getCellType() == CellType.FORMULA) {
                                    value = "=" + value;
                                }
                                out.print(encodeValue(value));
                            }
                            firstCell = false;
                        }
                        out.println();
                    }
                }
            }
            wb.close();
        } catch (Exception e) {
            System.out.print("eae");
            e.printStackTrace();
        }
    }
}