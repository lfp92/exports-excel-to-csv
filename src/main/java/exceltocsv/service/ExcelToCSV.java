package exceltocsv.service;

import exceltocsv.strategy.SheetTypeEnum;
import org.apache.poi.ss.usermodel.*;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintStream;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import static java.nio.charset.StandardCharsets.UTF_8;
import static org.apache.poi.ss.usermodel.Row.MissingCellPolicy.RETURN_BLANK_AS_NULL;

public class ExcelToCSV {

    private static final byte[] BYTE_ORDER_MARK = new byte[]{(byte) 0xEF, (byte) 0xBB, (byte) 0xBF};

    private static final Pattern REGEX_QUOTE = Pattern.compile("\"");

    private PrintStream printStream;

    private FormulaEvaluator formulaEvaluator;

    public void exportCSV(String inputFilePath, String outputFilePath) {

        try {
            Workbook workbook = getWorkbook(inputFilePath);

            formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();

            for (int sheetNo = 0, ns = workbook.getNumberOfSheets(); sheetNo < ns; sheetNo++) {

                Sheet sheet = workbook.getSheetAt(sheetNo);

                printStream = getPrintStream(outputFilePath, sheet.getSheetName());

                parseSheet(sheet);
            }

            workbook.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private Workbook getWorkbook(String inputFilePath) throws IOException {
        String fileExtension = getFileExtension(inputFilePath);

        return SheetTypeEnum.valueOf(fileExtension.toUpperCase())
                .getSheetType().getWorkbook(inputFilePath);
    }


    private void parseSheet(Sheet sheet) throws IOException {

        printStream.write(BYTE_ORDER_MARK);

        for (int rowCount = 0, lastRowNum = sheet.getLastRowNum(); rowCount <= lastRowNum; rowCount++) {

            parseRow(sheet.getRow(rowCount));
        }
    }

    private void parseRow(Row row) {
        if (isNotNull(row)) {
            for (int cellIndex = 0, cellNum = row.getLastCellNum(); cellIndex < cellNum; cellIndex++)
                parseCell(row.getCell(cellIndex, RETURN_BLANK_AS_NULL));

            printStream.println();
        } else {
            printStream.println(',');
        }
    }

    private void parseCell(Cell cell) {
        DataFormatter dataFormatter = new DataFormatter();

        if (isNotFirstCell(cell))
            printStream.print(',');

        if (isNotNull(formulaEvaluator))
            cell = formulaEvaluator.evaluateInCell(cell);

        String value = dataFormatter.formatCellValue(cell);

        if (isFormula(cell))
            value = String.format("=%s", value);

        printStream.print(escapeCharacters(value));
    }

    private String getFileExtension(String inputFilePath) {
        String[] arr = inputFilePath.split("(\\.)");
        return arr[arr.length - 1];
    }

    private PrintStream getPrintStream(String outputFilePath, String sheetName) throws FileNotFoundException {
        return new PrintStream(new FileOutputStream(outputFilePath + "_" + sheetName + ".csv"), true, UTF_8);
    }

    private boolean isNotFirstCell(Cell cell) {
        return cell.getColumnIndex() > 0;
    }

    private boolean isFormula(Cell cell) {
        return CellType.FORMULA.equals(cell.getCellType());
    }

    private boolean isNotNull(Object object) {
        return object != null;
    }

    private String escapeCharacters(String value) {

        Matcher matcher = REGEX_QUOTE.matcher(value);

        value = matcher.replaceAll("\"\"");

        if (isEscapableCharacter(value) || matcher.find())
            return "\"" + value + "\"";
        else
            return value;
    }

    private boolean isEscapableCharacter(String value) {
        return value.indexOf(',') != -1 || value.indexOf('"') != -1
                || value.indexOf('\n') != -1 || value.indexOf('\r') != -1;
    }
}
