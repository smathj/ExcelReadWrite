package excel;

import java.io.FileInputStream;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelFile {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		
		try {
			
			FileInputStream file = new FileInputStream("C:\\Users\\KSJ\\Desktop\\2020.xlsx");
			XSSFWorkbook wb = new XSSFWorkbook(file);
			int totalSheetNumber = wb.getNumberOfSheets(); 
			
			for(int s = 0; s< totalSheetNumber; s++) {	// 시트
			
				XSSFSheet sheet = wb.getSheetAt(s); 
				int rows = sheet.getPhysicalNumberOfRows(); 
				
				for(int i = 0; i< rows; i++) {			// 행
					
					XSSFRow row = sheet.getRow(i);	
					int cells = row.getPhysicalNumberOfCells(); 
					
					for(int j = 0; j < cells; j++) {	// 열
						
						if( row.getCell(j) == null) {
							continue;
						}
						System.out.print(row.getCell(j));
					}
					System.out.println();
				}
			}
			file.close();
		} catch (Exception e) {
			e.printStackTrace();
		} 
	}

}
