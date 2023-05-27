package excle_sheet;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class mock1 {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		
		FileInputStream file = new FileInputStream("F:\\software testing\\apache foi\\Book1.xlsx");
		
		Sheet sh = WorkbookFactory.create(file).getSheet("sheet3");
		
		int lastrowindex = sh.getLastRowNum();
		
		for(int i=0; i<=lastrowindex; i++)
		{
			int lastcellindex = sh.getRow(i).getLastCellNum()-1;
			 for(int j=0; j<=lastcellindex; j++)
			 {
				Cell cellinfo = sh.getRow(i).getCell(j);
				
				CellType ct = cellinfo.getCellType();
				
				if(ct==CellType.STRING)
				{
					String value = cellinfo.getStringCellValue();
					System.out.print(value+" | ");
				}
				else if (ct==CellType.NUMERIC) 
				{
					double value = cellinfo.getNumericCellValue();
					System.out.print(value+" | ");
				}
				else if (ct==CellType.BOOLEAN) 
				{
					boolean value = cellinfo.getBooleanCellValue();
					System.out.print(value+" | ");
				}
				}
				System.out.println(); 
				 
		}
			
			
	}
	
	
	
	
}
