package program;

import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

public class Excel_Writer {

	public static void main(String[] args){
		
		WebDriver driver;
		
		String path = ".//src//main//resources//data.xlsx";
		String url = "https://www.facebook.com";
		
		ChromeOptions option =  new ChromeOptions();
		option.setBinary("C:\\Program Files (x86)\\Google\\Chrome Beta\\Application\\chrome.exe");
		System.setProperty("webdriver.chrome.driver", "E:\\Selenium\\Drivers\\chromedriver.exe");
		
		driver = (WebDriver) new ChromeDriver(option);
				
		try{
			
				File file = new File(path);
				FileInputStream fs = new FileInputStream(file);
				
				Workbook wb = new XSSFWorkbook(fs);
			
				Sheet sheet = wb.getSheet("Sheet1");
				
				int rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();
				
				for(int i = 0; i < rowCount; i ++){
					
					Row row = sheet.getRow(i);
					
					for(int j = 0; j < row.getLastCellNum(); j ++){
						
						if(row.getCell(j).getCellType() == Cell.CELL_TYPE_NUMERIC)
							System.out.print(row.getCell(j).getNumericCellValue()+"||");
						
						else 
							System.out.print(row.getCell(j).getStringCellValue()+"||");
						
					}
					
					System.out.println();
				}
		} //End of try block
		
		catch(Exception e){
			
			System.out.println(e.getMessage());
		}
		
	}

}
