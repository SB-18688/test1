import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DayOne {
	public static void main(String[] args) throws IOException {
		File file = new File("C:\\Users\\Sathish Babu\\eclipse-workspace\\dayone\\Excel\\maven-sample1.xlsx");
		FileInputStream stream = new FileInputStream(file);
		Workbook workbook = new XSSFWorkbook(stream);
		Sheet sheet = workbook.getSheet("Sheet1");
//		Row row = sheet.getRow(1);
//		Cell cell = row.getCell(0);
////	System.out.println(cell);
		for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
			Row row = sheet.getRow(i);
			System.out.println("\nRow : "+i);
			for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
				Cell cell = row.getCell(j);
				int cellType = cell.getCellType();
				if (cellType == 1) {
					String string = cell.getStringCellValue();
					System.out.println(string);
				}
				if (cellType == 0) {
					if (DateUtil.isCellDateFormatted(cell)) {
						String format = new SimpleDateFormat("dd-mm-yy").format(cell.getDateCellValue());
						System.out.println(format);
					}

					else {
						double value = cell.getNumericCellValue();
						System.out.println(String.valueOf((long) value));
					}

				}

			}
		}

	}
}
