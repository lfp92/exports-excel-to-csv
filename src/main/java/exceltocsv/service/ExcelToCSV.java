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
    private static Pattern REGEX_QUOTE = Pattern.compile("\"");
    private String fileName;

    private static String encodeValue(String value) {
        boolean needQuotes = false;
        if (value.indexOf(',') != -1 || value.indexOf('"') != -1 || value.indexOf('\n') != -1
                || value.indexOf('\r') != -1)
            needQuotes = true;
        Matcher m = REGEX_QUOTE.matcher(value);
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
        String fileExtension = arr[arr.length - 1];
        Workbook workbook = null;
        PrintStream printStream = null;
        DataFormatter dataFormatter = new DataFormatter();
        try {

            byte[] bom = {(byte) 0xEF, (byte) 0xBB, (byte) 0xBF};
            FileInputStream file = new FileInputStream(new File(this.fileName));
            if (fileExtension.equals("xlsx")) {
                workbook = new XSSFWorkbook(file);
            } else {
                workbook = new HSSFWorkbook(new POIFSFileSystem(new File(this.fileName)));
            }

            FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();

            for (int sheetNo = 0, ns = workbook.getNumberOfSheets(); sheetNo < ns; sheetNo++) {
                Sheet sheet = workbook.getSheetAt(sheetNo);

                printStream = new PrintStream(new FileOutputStream(new File(outputFilePath + "_" + sheet.getSheetName() + ".csv")), true, "UTF-8");
                printStream.write(bom);
                for (int rowCount = 0, lastRowNum = sheet.getLastRowNum(); rowCount <= lastRowNum; rowCount++) {
                    Row row = sheet.getRow(rowCount);
                    if (row == null) {
                        printStream.println(',');
                        continue;
                    }
                    boolean firstCell = true;
                    for (int c = 0, cn = row.getLastCellNum(); c < cn; c++) {
                        Cell cell = row.getCell(c, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                        if (!firstCell)
                            printStream.print(',');
                        if (cell != null) {
                            if (formulaEvaluator != null)
                                cell = formulaEvaluator.evaluateInCell(cell);
                            String value = dataFormatter.formatCellValue(cell);
                            if (cell.getCellType() == CellType.FORMULA) {
                                value = "=" + value;
                            }
                            printStream.print(encodeValue(value));
                        }
                        firstCell = false;
                    }
                    printStream.println();
                }
            }

            workbook.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}