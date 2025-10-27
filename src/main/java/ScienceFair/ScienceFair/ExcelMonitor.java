package ScienceFair.ScienceFair;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelMonitor {

    public static void main(String[] args) {
    	// String filePath = "C:\\Users\\User\\eclipse-workspace\\ScienceFair\\src\\main\\java\\ScienceFair\\ScienceFair\\ScienceFair.xlsx"; 
    	// Path to the Excel file 
    	
    	File file = new File("src/main/java/ScienceFair/ScienceFair/ScienceFair.xlsx");
  	
        try (FileInputStream fis = new FileInputStream(file);
            Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0); // Assume the first sheet

            boolean errorsFound = false; 

            // Iterate over rows (skip the header row)
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                
                if (row == null) continue; // Skip empty rows

                Cell nameCell    = row.getCell(0);
                Cell projectCell = row.getCell(1);
                Cell judge1Cell  = row.getCell(2);
                Cell judge2Cell  = row.getCell(3);
                Cell avgCell     = row.getCell(4);
                
                // Check for missing data 
                if (isCellEmpty(nameCell) || isCellEmpty(projectCell)) {
                	System.out.println("Row " + (i + 1) + ": Missing student name or project title."); 
                    errorsFound = true; 
                } 
                
                // Check score validity 
                double judge1 = getCellNumericValue(judge1Cell); 
                double judge2 = getCellNumericValue(judge2Cell); 
                
                if (judge1 < 0 || judge1 > 10) {
                	System.out.println("Row " + (i + 1) + ": Judge 1 score is out of range."); 
                    errorsFound = true; 
                }
                
                if (judge2 < 0 || judge2 > 10) {
                	System.out.println("Row " + (i + 1) + ": Judge 2 score is out of range."); 
                    errorsFound = true; 
                } 
                
                // Check average score calculation 
                if (!isCellEmpty(avgCell) && judge1 >= 0 && judge2 >= 0) { 
                    double correctAvg = Math.round((judge1 + judge2) / 2 * 10.0) / 10.0; 
                    double avg = getCellNumericValue(avgCell);
                    
                    if (avg != correctAvg) {
                    	System.out.println("Row " + (i + 1) + ": Incorrect average score. Expected: " + correctAvg); 
                        errorsFound = true; 
                    }
                }
            }
            
            if (!errorsFound) {
            	System.out.println("All data is correct!"); 
            } 
        } catch (IOException e) {
            System.err.println("Error reading the Excel file: " + e.getMessage());
        }
    }

    private static boolean isCellEmpty(Cell cell) {
    	return cell == null || cell.getCellType() == CellType.BLANK; 
    } 

    private static double getCellNumericValue(Cell cell) { 
        if (cell == null || cell.getCellType() != CellType.NUMERIC) { 
            return -1; // Return -1 to indicate invalid or missing numeric data 
        } 
        
        return cell.getNumericCellValue(); 
    }
}
