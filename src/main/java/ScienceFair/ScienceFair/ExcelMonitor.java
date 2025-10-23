package ScienceFair.ScienceFair;

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
        // change this to your actual file, or pass it as arg[0]
        String filePath = args.length > 0
                ? args[0]
                : "C:\\Users\\User\\eclipse-workspace\\ScienceFair\\src\\main\\java\\ScienceFair\\ScienceFair\\ScienceFair.xlsx";

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0); // first sheet
            int last = sheet.getLastRowNum();

            System.out.println("Row\tName\tProject\tJudge1\tJudge2\tAvg");
            for (int i = 1; i <= last; i++) { // assume row 0 is header
                Row row = sheet.getRow(i);
                if (row == null) continue;

                Cell nameCell    = row.getCell(0);
                Cell projectCell = row.getCell(1);
                Cell judge1Cell  = row.getCell(2);
                Cell judge2Cell  = row.getCell(3);
                Cell avgCell     = row.getCell(4);

                String name    = getString(nameCell);
                String project = getString(projectCell);

                Double j1 = getNumeric(judge1Cell);
                Double j2 = getNumeric(judge2Cell);
                Double avg = null;
                if (j1 != null && j2 != null) {
                    avg = (j1 + j2) / 2.0;
                } else if (!isCellEmpty(avgCell) && avgCell.getCellType() == CellType.NUMERIC) {
                    // if file already has an average
                    avg = avgCell.getNumericCellValue();
                }

                System.out.printf("%d\t%s\t%s\t%s\t%s\t%s%n",
                        i,
                        name,
                        project,
                        j1 == null ? "NA" : j1,
                        j2 == null ? "NA" : j2,
                        avg == null ? "NA" : String.format("%.2f", avg));
            }

        } catch (IOException e) {
            System.err.println("Error reading the Excel file: " + e.getMessage());
        }
    }

    private static boolean isCellEmpty(Cell cell) {
        return cell == null || cell.getCellType() == CellType.BLANK
               || (cell.getCellType() == CellType.STRING
                   && cell.getStringCellValue() != null
                   && cell.getStringCellValue().trim().isEmpty());
    }

    private static String getString(Cell cell) {
        if (cell == null) return "";
        if (cell.getCellType() == CellType.STRING) return cell.getStringCellValue();
        if (cell.getCellType() == CellType.NUMERIC) return String.valueOf(cell.getNumericCellValue());
        if (cell.getCellType() == CellType.BOOLEAN) return String.valueOf(cell.getBooleanCellValue());
        return "";
    }

    private static Double getNumeric(Cell cell) {
        if (cell == null) return null;
        if (cell.getCellType() == CellType.NUMERIC) return cell.getNumericCellValue();
        // also allow numeric-looking strings
        if (cell.getCellType() == CellType.STRING) {
            try {
                return Double.parseDouble(cell.getStringCellValue().trim());
            } catch (NumberFormatException ignored) { }
        }
        return null;
    }
}
