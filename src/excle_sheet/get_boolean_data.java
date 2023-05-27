package excle_sheet;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class get_boolean_data {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		
		FileInputStream file = new FileInputStream("F:\\software testing\\apache foi\\Book1.xlsx");
		
		boolean value = WorkbookFactory.create(file).getSheet("sheet1").getRow(1).getCell(1).getBooleanCellValue();
		
		System.out.println(value);
		
		
		
		
		
	}
	
	
	
	
}
