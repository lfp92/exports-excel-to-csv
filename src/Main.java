import exceltocsv.ExcelToCSV;

public class Main {
	public static void main(String[] args) {
		String fileName = "C:\\Users\\leonardo.petrauskas\\Downloads\\Pasta1.xls";
		String output = "C:\\Users\\leonardo.petrauskas\\Downloads\\teste";
		ExcelToCSV ex = new ExcelToCSV();
		ex.exportCSV(fileName, output);
	}
}