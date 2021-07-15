package doc1;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Doc1_Reader {
	
	
private static final String NAME = "src\\main\\java\\data\\DoesFileExist.xlsx";

	
	public static void main(String[] args) {
		try {

			FileInputStream file = new FileInputStream(new File(NAME));
			Workbook workbook = new XSSFWorkbook(file);
			DataFormatter dataFormatter = new DataFormatter();
			Iterator<Sheet> sheets = workbook.sheetIterator();
			while (sheets.hasNext()) {
				Sheet sh = sheets.next();
				System.out.println("Sheet name is " + sh.getSheetName());
				Iterator<Row> iterator = sh.iterator();
				while (iterator.hasNext()) {
					Row row = iterator.next();
					Iterator<Cell> celliterator = row.iterator();
					while (celliterator.hasNext()) {
						Cell cell = celliterator.next();
						String cellValue = dataFormatter.formatCellValue(cell);

						System.out.println(cellValue + "\t");
					}
					System.out.println();

				}

			}

			workbook.close();

		} catch (Exception e) {

			e.printStackTrace();
		}

	}

}

