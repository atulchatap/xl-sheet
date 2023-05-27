package excle_sheet;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.logging.log4j.util.FilteredObjectInputStream;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class get_rowsize_inasheet {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		
		FileInputStream file= new FileInputStream("F:\\software testing\\apache foi\\Book1.xlsx");
		
		int row = WorkbookFactory.create(file).getSheet("sheet1").getLastRowNum()+1;
		
		System.out.println(row);
		
		
		
	}
}
