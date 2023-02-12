package StylingLaporan;

import org.apache.poi.ss.usermodel.*;
import org.jetbrains.annotations.NotNull;

public class StylerRepo {

    public void centerTextStyle(@NotNull Workbook BookPertindakanNew) {
        CellStyle centerTextStyle = BookPertindakanNew.createCellStyle();
        centerTextStyle.setAlignment(HorizontalAlignment.CENTER);
    }

    public void AllBorderStyle(@NotNull Workbook BookPertindakanNew) {
        CellStyle AllBorderStyle = BookPertindakanNew.createCellStyle();
        AllBorderStyle.setBorderBottom(BorderStyle.THIN);
        AllBorderStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        AllBorderStyle.setBorderLeft(BorderStyle.THIN);
        AllBorderStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        AllBorderStyle.setBorderRight(BorderStyle.THIN);
        AllBorderStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
        AllBorderStyle.setBorderTop(BorderStyle.THIN);
        AllBorderStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
    }
}
