package excle_sheet;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class get_String_data {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		
		FileInputStream file= new FileInputStream("F:\\software testing\\apache foi\\Book1.xlsx");
		
	String value = WorkbookFactory.create(file).getSheet("sheet1").getRow(1).getCell(0).getStringCellValue();
		
		System.out.println(value);
		
		
		
	}
	
	
	
}
