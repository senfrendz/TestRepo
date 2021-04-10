import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class excelOps {
	
	static File source = new File("ExcelFile\\test.xlsx");

	
	public static void excelReader() throws IOException {
		
		try {
			FileInputStream fileReader = new FileInputStream(source);
			
			@SuppressWarnings("resource")
			XSSFWorkbook wb = new XSSFWorkbook(fileReader);
			
			XSSFSheet sheet = wb.getSheetAt(0);
			
			int count = sheet.getLastRowNum();
			
			
			for(int i=0;i<=count;i++) {
				
				System.out.println(sheet.getRow(i).getCell(0).getStringCellValue() +" | "+ sheet.getRow(i).getCell(1).getStringCellValue());
			}
						
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}

	public static void excelWriter() throws IOException {
		
		FileInputStream src = new FileInputStream(source);
		
		XSSFWorkbook workbook = new XSSFWorkbook(src);
		
		XSSFSheet testSheet = workbook.getSheetAt(0);
		
		int rowCount = testSheet.getLastRowNum();
		
		for(int j=1;j<rowCount;j++) {
			
			if(j%2==0) {
			
			 testSheet.getRow(j).createCell(3).setCellValue("Pass value");
			}
			else {
				 testSheet.getRow(j).createCell(3).setCellValue("Fail value");				
			}
		}
		
		FileOutputStream fout = new FileOutputStream("ExcelFile\\sen.xlsx");
		workbook.write(fout);
		
	}
	
	public static void main(String[] args) throws IOException {
	//	excelReader();
		excelWriter();
	}

}
