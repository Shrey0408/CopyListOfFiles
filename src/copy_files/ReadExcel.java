package copy_files;

import java.io.File;  
import java.io.FileInputStream;  
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.hssf.usermodel.HSSFSheet;  
import org.apache.poi.hssf.usermodel.HSSFWorkbook;  
import org.apache.poi.ss.usermodel.Cell;  
import org.apache.poi.ss.usermodel.FormulaEvaluator;  
import org.apache.poi.ss.usermodel.Row;  
/*
 * Excel format
 * Dest.Folder name, source path, file name
 * 
 */
public class ReadExcel  
{  
	@SuppressWarnings("deprecation")
	public static List<ArrayList<String>> getExcelData(String path) throws IOException  
	{  

		FileInputStream fis=new FileInputStream(new File(path));  
		String cellData = null;
		List<ArrayList<String>> data = new ArrayList<>();

		try (HSSFWorkbook wb = new HSSFWorkbook(fis)) {
			HSSFSheet sheet=wb.getSheetAt(0);  
			FormulaEvaluator formulaEvaluator=wb.getCreationHelper().createFormulaEvaluator();  
			for(Row row: sheet)     
			{  
				ArrayList<String> items = new ArrayList<>();
				for(Cell cell: row)    
				{  
				
					switch(formulaEvaluator.evaluateInCell(cell).getCellType())  
					{  
					case Cell.CELL_TYPE_NUMERIC:   
						cellData = ""+cell.getNumericCellValue();
						System.out.print(cellData+ "\t\t");
						items.add(cellData);
						break;  
					case Cell.CELL_TYPE_STRING:
						cellData = ""+cell.getStringCellValue();
						System.out.print(cellData+ "\t\t");
						items.add(cellData);
						System.out.print(cellData+ "\t\t");  
						break;  
					}  
					
				}  
				data.add(items);
				System.out.println();  
			}
		}
		return data;  
	}  
}  