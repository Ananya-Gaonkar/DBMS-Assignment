
import java.io.File;
import java.io.FileOutputStream;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import org.apache.poi.ss.usermodel.Row;  
import org.apache.poi.ss.usermodel.Cell;  

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelFileWriting {

	public static void main(String[] args) {
		XSSFWorkbook workbook = new XSSFWorkbook(); 
		XSSFSheet sheet = workbook.createSheet("Employee Data");


		Map<String, Object[]> data = new TreeMap<String, Object[]>();
		data.put("1", new Object[] {"USN", "NAME"});
		data.put("2", new Object[] {1, "Abdul"});
		data.put("3", new Object[] {38, "Joy"});
		data.put("4", new Object[] {63, "Pratham"});

		Set<String> keyset = data.keySet();
		int rownum = 0;
		for (String key : keyset) {

		  Row row = sheet.createRow(rownum++);
		  Object [] objArr = data.get(key);
		  int cellnum = 0;
		  for (Object obj : objArr)
		  {
		     Cell cell = row.createCell(cellnum++);
		     if(obj instanceof String)
		          cell.setCellValue((String)obj);
		      else if(obj instanceof Integer)
		          cell.setCellValue((Integer)obj);
		  }
		}

		try {
		  FileOutputStream out = new FileOutputStream(new File("output.xlsx"));
		  workbook.write(out);
		  out.close();
		  System.out.println("howtodoinjava_demo.xlsx written successfully on disk.");
		} 
		catch (Exception e) {
		  e.printStackTrace();
		}
	}

}