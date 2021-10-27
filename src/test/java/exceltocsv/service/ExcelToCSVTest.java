package exceltocsv.service;

import lombok.SneakyThrows;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import java.io.File;

class ExcelToCSVTest {

    private ExcelToCSV excelToCSV = new ExcelToCSV();

    @Test
    @SneakyThrows
    void deveExtrairPlanilhas() {
        excelToCSV.exportCSV("C:\\Users\\Leonardo\\Downloads\\ficha-teste.xlsx", "C:\\Users\\Leonardo\\Downloads\\teste");

        Assertions.assertDoesNotThrow(() -> {
            File fileA = new File("C:\\Users\\Leonardo\\Downloads\\teste_Ficha A.csv");
            File fileB = new File("C:\\Users\\Leonardo\\Downloads\\teste_Ficha B.csv");
        });

    }

}