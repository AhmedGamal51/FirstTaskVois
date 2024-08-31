import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.sikuli.script.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.time.temporal.ChronoUnit;

class Main {
    public static void main(String[] args) throws IOException, InterruptedException {
        String filePath = "C:\\Users\\Dell\\Desktop\\Vois_task\\src\\main\\resources\\TaskData.xlsx";
        FileInputStream fis = new FileInputStream(new File(filePath));
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheetAt(0);

        int joiningDateColumnIndex = 2; // "Joining Date" is in column C (index 2)
        int yearsSpentColumnIndex = 3; // "Years Spent" is in column D (index 3)

        // Update the formatter to ignore the day of the week and only parse the date part
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("MMMM dd, yyyy");
        LocalDate currentDate = LocalDate.now();

        for (Row row : sheet) {
            if (row.getRowNum() == 0) continue; // Skip header row

            Cell joiningDateCell = row.getCell(joiningDateColumnIndex);

            if (joiningDateCell != null && joiningDateCell.getCellType() == CellType.STRING) {
                String joiningDateStr = joiningDateCell.getStringCellValue().trim();

                // Remove the day of the week portion ("Friday, ", etc.) before parsing
                if (!joiningDateStr.isEmpty()) {
                    try {
                        // Remove the first part of the string before the comma (day of the week)
                        int commaIndex = joiningDateStr.indexOf(",");
                        if (commaIndex != -1) {
                            joiningDateStr = joiningDateStr.substring(commaIndex + 1).trim();
                        }

                        // Parse the date with the new format
                        LocalDate joiningDate = LocalDate.parse(joiningDateStr, formatter);
                        long yearsSpent = ChronoUnit.YEARS.between(joiningDate, currentDate);

                        Cell yearsSpentCell = row.createCell(yearsSpentColumnIndex);
                        yearsSpentCell.setCellValue(yearsSpent);
                    } catch (DateTimeParseException e) {
                        System.out.println("Skipping invalid date format at row " + (row.getRowNum() + 1) + ": " + joiningDateStr);
                    }
                }
            }
        }

//         Save the modified Excel file
        fis.close();
        FileOutputStream fos = new FileOutputStream(new File(filePath));
        workbook.write(fos);
        fos.close();
        workbook.close();

        System.out.println("Excel sheet updated successfully!");


    }
}


