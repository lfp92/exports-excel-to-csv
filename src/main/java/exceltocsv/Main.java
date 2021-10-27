package exceltocsv;

import exceltocsv.service.ExcelToCSV;

public class Main {
    public static void main(String[] args) {
        String fileName = "C:\\Users\\Leonardo\\Downloads\\ficha-teste.xlsx";
        String output = "C:\\Users\\Leonardo\\Downloads\\teste";
        ExcelToCSV ex = new ExcelToCSV();
        ex.exportCSV(fileName, output);
    }
}