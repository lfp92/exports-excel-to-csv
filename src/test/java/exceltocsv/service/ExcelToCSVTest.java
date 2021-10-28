package exceltocsv.service;

import lombok.SneakyThrows;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import java.io.File;

class ExcelToCSVTest {

    public static final String PATH = "C:\\Users\\Leonardo\\Downloads\\";
    private final ExcelToCSV excelToCSV = new ExcelToCSV();

    @Test
    @SneakyThrows
    void deveExtrairPlanilhas() {
        excelToCSV.exportCSV(PATH.concat("ficha-teste.xlsx"), PATH.concat("teste"));

        File fileA = new File(PATH.concat("teste_Ficha A.csv"));
        File fileB = new File(PATH.concat("teste_Ficha B.csv"));

        Assertions.assertTrue(fileA.isFile());
        Assertions.assertTrue(fileB.isFile());

    }

}