package excle_sheet;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class get_cellsize_inasheet {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		
		FileInputStream file =new FileInputStream("F:\\software testing\\apache foi\\Book1.xlsx");
		
		int lastcell = WorkbookFactory.create(file).getSheet("sheet1").getRow(0).getLastCellNum();
		
		System.out.println(lastcell);
		
	}
}
