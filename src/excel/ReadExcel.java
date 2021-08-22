package excel;



import java.io.FileInputStream;
import java.io.IOException;

import javax.servlet.Servlet;
import javax.servlet.ServletConfig;
import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Servlet implementation class test
 */
@WebServlet("/readExcel")
public class ReadExcel extends HttpServlet {
	private static final long serialVersionUID = 1L;

    /**
     * Default constructor. 
     */
    public ReadExcel() {
        // TODO Auto-generated constructor stub
    }

	/**
	 * @see Servlet#init(ServletConfig)
	 */
	public void init(ServletConfig config) throws ServletException {
		// TODO Auto-generated method stub
	}

	/**
	 * @see Servlet#destroy()
	 */
	public void destroy() {
		// TODO Auto-generated method stub
	}

	/**
	 * @see HttpServlet#doGet(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doGet(HttpServletRequest request, HttpServletResponse response) {
		// TODO Auto-generated method stub
		
		// ���� �б�
		
		
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
