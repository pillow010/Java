import java.util.HashSet;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.jetbrains.annotations.NotNull;

public class RemoveDuplicates {
    public static void removeDuplicates(@NotNull Sheet sheet) {
        Set<String> uniqueRows = new HashSet<>();
        int lastRowNum = sheet.getLastRowNum();
        for (int i = 0; i <= lastRowNum; i++) {
            Row currentRow = sheet.getRow(i);
            StringBuilder sb = new StringBuilder();
            for (int j = 0; j < currentRow.getLastCellNum(); j++) {
                Cell currentCell = currentRow.getCell(j, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                switch (currentCell.getCellType()) {
                    case STRING:
                        sb.append(currentCell.getStringCellValue());
                        break;
                    case NUMERIC:
                        sb.append(currentCell.getNumericCellValue());
                        break;
                }
            }
            String rowAsString = sb.toString();
            if (uniqueRows.contains(rowAsString)) {
                sheet.removeRow(currentRow);
                i--;
                lastRowNum--;
            } else {
                uniqueRows.add(rowAsString);
            }
        }
    }
}
