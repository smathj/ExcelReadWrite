package excel;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateExcel {
	
	
	
	@SuppressWarnings({ "unused", "resource" })
	public static void main(String[] args) throws Exception {
	
	
        //workbook = new HSSFWorkbook(); // 엑셀 97 ~ 2003
        //workbook = new XSSFWorkbook(); // 엑셀 2007 버전 이상
		
			
			XSSFWorkbook workbook = new XSSFWorkbook();
			XSSFSheet sheet = null;
			XSSFRow row = null;
			XSSFCell cell = null;
			
			int maxSheet = 10;
			int maxCell = 5;
			int maxRow = 5;
			
	        //XSSFSheet sheet = workbook.createSheet("첫번쨰 시트"); // 새 시트(Sheet) 생성
		   
	        //XSSFRow row = sheet.createRow(0); 				// 엑셀의 행은 0번부터 시작
		    
	        //XSSFCell cell = row.createCell(0); 				// 행의 셀은 0번부터 시작
		    
	        //cell.setCellValue("테스트 데이터"); 				//생성한 셀에 데이터 삽입
	        
	        for(int i =  0; i < maxSheet; i++) {
	        	sheet = workbook.createSheet((i+1) + "번 시트");  // n번째 시트 생성
	        	
	        	
	        	for(int r = 0; r < maxRow; r++) {
	        		
	        		row = sheet.createRow(r); 				  		// n번 로우 생성
	        		
	        		for(int c = 0; c < maxCell; c++) {
	        			cell = row.createCell(c);					// n로우 , n번쨰 셀 생성
        				cell.setCellValue((c+1)*(r+1) + "번째");
	        		}
	        	}
	        }
	        
	        
		    
		    
		    
	        try {
	            FileOutputStream fileoutputstream = new FileOutputStream("C:\\Users\\KSJ\\Desktop\\만든엑셀.xlsx");
	            workbook.write(fileoutputstream);
	            fileoutputstream.close();
	            System.out.println("엑셀파일생성성공");
	            
	        } catch (IOException e) {
	            e.printStackTrace();
	            System.out.println("엑셀파일생성실패");
	            
	        }


			
			
		
		
		
	}
}
