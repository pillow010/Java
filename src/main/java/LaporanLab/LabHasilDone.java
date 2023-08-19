package LaporanLab;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.regex.Pattern;

public class LabHasilDone {
    public static void main(String[] args) {
        Workbook bookHasilRinci = null;
        XSSFWorkbook newSheetBook = new XSSFWorkbook();
        FileOutputStream outputStream = null;
        boolean doneFinal = true;

        String localDate = LocalDate.now ().minusMonths (1).format (DateTimeFormatter.ofPattern ("yy MM"));
//        DateTimeFormatter formatter = DateTimeFormatter.ofPattern ("yyyyMMdd HHmmss");
//        String formattedDateTime = LocalDateTime.now ().format (formatter);
        Pattern pattern = Pattern.compile("[\\\\/:*?\"<>|]"); // Invalid characters for sheet names
        String fileInput = "C:\\sat work\\test\\1. input\\";
        String fileOutput = "C:\\sat work\\test\\2. output\\";
        String fileNameHasilRinci = localDate + " lab hasil rinci";
        String fileNameOutputDone      = fileOutput + "Done Lab Hasil " + localDate + ".xlsx";
        String fileNameOutputHalfDone = fileOutput + fileNameHasilRinci + " half done.xlsx";
        String[] pemeriksaanDicari ={
                "Anti HAV IgG/IgM", "Anti HCV (Rapid)", "Anti HIV", "CD4 Paket", "HAV Total", "HBS Ag (MCU)", "HBsAg",
                "HIV 1 & HIV 2", "WIDAL"
        };


//      Make Styling (allBorder for content and borderCenter for title)
        CellStyle AllBorderCellStyle = newSheetBook.createCellStyle ();
        AllBorderCellStyle.setBorderBottom (BorderStyle.THIN);
        AllBorderCellStyle.setBottomBorderColor (IndexedColors.BLACK.getIndex ());
        AllBorderCellStyle.setBorderLeft (BorderStyle.THIN);
        AllBorderCellStyle.setLeftBorderColor (IndexedColors.BLACK.getIndex ());
        AllBorderCellStyle.setBorderRight (BorderStyle.THIN);
        AllBorderCellStyle.setRightBorderColor (IndexedColors.BLACK.getIndex ());
        AllBorderCellStyle.setBorderTop (BorderStyle.THIN);
        AllBorderCellStyle.setTopBorderColor (IndexedColors.BLACK.getIndex ());
        CellStyle BorderCenterCellStyle = newSheetBook.createCellStyle ();
        BorderCenterCellStyle.setAlignment (HorizontalAlignment.CENTER);
        BorderCenterCellStyle.setBorderBottom (BorderStyle.THIN);
        BorderCenterCellStyle.setBottomBorderColor (IndexedColors.BLACK.getIndex ());
        BorderCenterCellStyle.setBorderLeft (BorderStyle.THIN);
        BorderCenterCellStyle.setLeftBorderColor (IndexedColors.BLACK.getIndex ());
        BorderCenterCellStyle.setBorderRight (BorderStyle.THIN);
        BorderCenterCellStyle.setRightBorderColor (IndexedColors.BLACK.getIndex ());
        BorderCenterCellStyle.setBorderTop (BorderStyle.THIN);
        BorderCenterCellStyle.setTopBorderColor (IndexedColors.BLACK.getIndex ());


        File xlsxHasilRinci = new File (fileInput  + fileNameHasilRinci   +".xlsx");
        File xlsHasilRinci  = new File (fileInput   + fileNameHasilRinci   +".xls");
        File fileHasilRinci;

        if (xlsxHasilRinci.exists ()) {
            fileHasilRinci = xlsxHasilRinci;
        } else if (xlsHasilRinci.exists ()) {
            fileHasilRinci = xlsHasilRinci;
        } else {
            System.out.println ("File not found: " + fileInput + fileNameHasilRinci);
            return;
        }


        try {
            InputStream hasilRinci = new FileInputStream(fileHasilRinci);
            bookHasilRinci = new XSSFWorkbook(hasilRinci);
            Sheet sheetHasilRinci = bookHasilRinci.getSheetAt(0);
            int lastRow = sheetHasilRinci.getLastRowNum();
            int lastCell = sheetHasilRinci.getRow(0).getLastCellNum();

            for (String pemeriksaan : pemeriksaanDicari) {
                String cleanedSheetName = pemeriksaan.replaceAll(pattern.pattern(), "");
                Sheet newSheet = newSheetBook.createSheet(cleanedSheetName);
                createTitleRow(sheetHasilRinci, newSheet, lastCell);

                for (int i = 1; i <= lastRow; i++) {
                    String cellValue = sheetHasilRinci.getRow(i).getCell(9).getStringCellValue();
                    if (cellValue.contains(pemeriksaan)) {
                        copyRow(sheetHasilRinci.getRow(i), newSheet.createRow(newSheet.getLastRowNum() + 1));
                    }
                }
            }

            // Loop through the sheets
            for (int i=0;i<newSheetBook.getNumberOfSheets ();i++) {
                // Loop through the cells in the first row
                Sheet doingSheet = newSheetBook.getSheetAt (i);
                System.out.println (doingSheet.getSheetName ());
                System.out.println (doingSheet.getLastRowNum ());
                for (int rightCell = 0; rightCell < doingSheet.getRow(0).getLastCellNum(); rightCell++) {
                    doingSheet.getRow(0).getCell(rightCell).setCellStyle(BorderCenterCellStyle);
                    doingSheet.autoSizeColumn(rightCell);
                    // Loop through the cells in the remaining rows
                    for (int downRow = 1; downRow <= doingSheet.getLastRowNum(); downRow++) {
                        if (doingSheet.getRow (downRow).getCell (rightCell) == null) {
                            doingSheet.getRow (downRow).createCell (rightCell).setCellValue ("");
                        } else {
                            doingSheet.getRow (downRow).getCell (rightCell).setCellStyle (AllBorderCellStyle);
                        }
                    }
                }
            }


            System.out.println (newSheetBook.getNumberOfSheets ());
        } catch (Exception e) {
            e.printStackTrace();
        }




        try {
            if (doneFinal) {
                outputStream = new FileOutputStream (fileNameOutputDone);
            } else {
                outputStream = new FileOutputStream (fileNameOutputHalfDone);
            }
            newSheetBook.write (outputStream);
            System.out.println ("file saved at "+fileOutput);
        } catch (
                IOException e) {
            e.printStackTrace ();
        } finally {
            try {
                if (bookHasilRinci != null) {
                    bookHasilRinci.close ();
                }
                if (outputStream != null) {
                    outputStream.close ();
                }
            } catch (IOException e) {
                e.printStackTrace ();
            }
        }
    }

    private static void createTitleRow(Sheet sourceSheet, Sheet targetSheet, int lastCell) {
        Row titleRow = targetSheet.createRow(0);
        for (int cll = 0; cll < lastCell; cll++) {
            Cell newCell = titleRow.createCell(cll);
            newCell.setCellValue(sourceSheet.getRow(0).getCell(cll).getStringCellValue());
        }
    }

    private static void copyRow(Row sourceRow, Row targetRow) {
        for (int j = 0; j < sourceRow.getLastCellNum(); j++) {
            Cell sourceCell = sourceRow.getCell(j);
            Cell targetCell = targetRow.createCell(j);

            if (sourceCell != null) {
                if (sourceCell.getCellType() == CellType.STRING) {
                    targetCell.setCellValue(sourceCell.getStringCellValue());
                } else {
                    targetCell.setCellValue(sourceCell.getNumericCellValue());
                }
            }
        }
    }

}
