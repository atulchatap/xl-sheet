package excle_sheet;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class assignment_get_alldata {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		
	FileInputStream file = new FileInputStream("F:\\software testing\\apache foi\\Book1.xlsx");
	
	 Sheet sh = WorkbookFactory.create(file).getSheet("sheet2");
		
		
		int lastrowindex = sh.getLastRowNum();
		
		for(int i=0; i<=lastrowindex; i++)
		{
			int lastcellindex = sh.getRow(i).getLastCellNum()-1;
			for(int j=0; j<=lastcellindex; j++)
			{
				double value = sh.getRow(i).getCell(j).getNumericCellValue();
				System.out.print(value+"  ");
			}
			System.out.println();
			
		}
		
		
		
		
	}
	
}
