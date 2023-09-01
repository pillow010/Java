package LaporanLab;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;
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
        Pattern pattern = Pattern.compile("[\\\\/:*?\"<>|.]"); // Invalid characters for sheet names
        String fileInput = "C:\\sat work\\test\\1. input\\";
        String fileOutput = "C:\\sat work\\test\\2. output\\";
        String fileNameHasilRinci = localDate + " lab hasil rinci";
        String fileNamePertindakanNew = localDate + " lab tindakan new";
        String fileNameOutputDone      = fileOutput + "Done Lab Hasil " + localDate + ".xlsx";
        String fileNameOutputHalfDone = fileOutput + fileNameHasilRinci + " half done.xlsx";
        String[] pemeriksaanDicari ={
                "Anti HAV IgG/IgM", "Anti HCV (Rapid)", "Anti HIV", "CD4 Paket", "HAV Total", "HBsAg", "HBsAg Final",
                "HIV 1 & HIV 2", "WIDAL", "WIDAL Final"
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
        File xlsxTindakanNew = new File (fileInput  + fileNamePertindakanNew   +".xlsx");
        File xlsTindakanNew  = new File (fileInput   + fileNamePertindakanNew  +".xls");
        File fileTindakanNew;

        if (xlsxHasilRinci.exists ()) {
            fileHasilRinci = xlsxHasilRinci;
        } else if (xlsHasilRinci.exists ()) {
            fileHasilRinci = xlsHasilRinci;
        } else {
            System.out.println ("File not found: " + fileInput + fileNameHasilRinci);
            return;
        }


        if (xlsxTindakanNew.exists ()) {
            fileTindakanNew = xlsxTindakanNew;
        } else if (xlsTindakanNew.exists ()) {
            fileTindakanNew = xlsTindakanNew;
        } else {
            System.out.println ("File not found: " + fileInput + fileNamePertindakanNew);
            return;
        }

        try {
            InputStream hasilRinci = new FileInputStream(fileHasilRinci);
            bookHasilRinci = WorkbookFactory.create (hasilRinci);
            FileInputStream inputStream = new FileInputStream(fileTindakanNew);
            Workbook bookPertindakan = WorkbookFactory.create(inputStream);
            Sheet sheetPerTindakan = bookPertindakan.getSheetAt (0);
            Sheet sheetHasilRinci = bookHasilRinci.getSheetAt(0);
            int lastCell = sheetHasilRinci.getRow(0).getLastCellNum();
            sheetHasilRinci.getRow (0).createCell (lastCell).setCellValue ("Diagnosa");

            // Create a HashMap to store the keys and corresponding values from sheetPerTindakan
            HashMap<String, String> perTindakanMap = new HashMap<> ();
            for (int j = 1; j <= sheetPerTindakan.getLastRowNum(); j++) {
                String key = sheetPerTindakan.getRow (j).getCell (0).getStringCellValue () +
                        sheetPerTindakan.getRow (j).getCell (1).getStringCellValue () +
                        sheetPerTindakan.getRow (j).getCell (2).getStringCellValue () +
                        sheetPerTindakan.getRow (j).getCell (3).getStringCellValue () +
                        sheetPerTindakan.getRow (j).getCell (4).getStringCellValue ();
                Cell cellValue = sheetPerTindakan.getRow (j).getCell (13);
                if (cellValue == null) {
                    perTindakanMap.put (key, "");
                } else {
                    String value = cellValue.getStringCellValue ();
                    perTindakanMap.put (key, value);
                }
            }

            // Iterate through sheetHasilRinci and perform lookups in the HashMap
            for (int i = 1; i <= sheetHasilRinci.getLastRowNum(); i++) {
                String noreg = sheetHasilRinci.getRow(i).getCell(0).getStringCellValue().replaceAll(pattern.pattern(), "");
                sheetHasilRinci.getRow(i).getCell(0).setCellValue(noreg);
                Cell cellHasil = sheetHasilRinci.getRow (i).getCell (11);
                Cell cellPemeriksaan = sheetHasilRinci.getRow (i).getCell (9);
                if (cellHasil.getStringCellValue ().contains ("/") && cellPemeriksaan.getStringCellValue ().equalsIgnoreCase ("widal")){
                    cellHasil.setCellValue ("POSITIVE");
                }
                String keyToLookup = sheetHasilRinci.getRow(i).getCell(0).getStringCellValue();
                String valueFromMap = perTindakanMap.get(keyToLookup);

                if (valueFromMap != null) {
                    sheetHasilRinci.getRow(i).createCell(lastCell).setCellValue(valueFromMap);
                }
            }

            int lastRow = sheetHasilRinci.getLastRowNum();
            int cellDiagnosa = sheetHasilRinci.getRow (0).getLastCellNum ();
            sheetHasilRinci.getRow (0).createCell (cellDiagnosa).setCellValue ("");
            lastCell = sheetHasilRinci.getRow(0).getLastCellNum();

            for (String pemeriksaan : pemeriksaanDicari) {
                String cleanedSheetName = pemeriksaan.replaceAll(pattern.pattern(), "");
                Sheet newSheet = newSheetBook.createSheet(cleanedSheetName);
                createTitleRow(sheetHasilRinci, newSheet, lastCell);

                for (int i = 1; i <= lastRow; i++) {
                    Row currentRow = sheetHasilRinci.getRow(i);
                    Cell cell9 = currentRow.getCell(9);
                    String cellValue = cell9.getStringCellValue();

                    if (cellValue.contains("HBS Ag")) {
                        String[] splitValue = cellValue.split("HBS Ag");
                        String HBSAgMcu = "HBsAg";
                        String lastChar = splitValue[1];
                        cell9.setCellValue(HBSAgMcu + lastChar);
                    }

                    if (cellValue.contains(pemeriksaan)) {
                        copyRow(currentRow, newSheet.createRow(newSheet.getLastRowNum() + 1));
                    }
                }
                System.out.println("Sheet " + pemeriksaan + " filed");
            }

            // Loop through the sheets
            for (int i = 0; i < newSheetBook.getNumberOfSheets(); i++) {
                // Loop through the cells in the first row
                Sheet doingSheet = newSheetBook.getSheetAt(i);
//                System.out.println(doingSheet.getSheetName() + " tidied");
//                System.out.println(doingSheet.getLastRowNum());
                for (int rightCell = 0; rightCell < lastCell; rightCell++) {
                    doingSheet.getRow(0).getCell(rightCell).setCellStyle(BorderCenterCellStyle);
                    // Loop through the cells in the remaining rows
                    for (int downRow = 1; downRow <= doingSheet.getLastRowNum(); downRow++) {
                        Row currentRow = doingSheet.getRow(downRow);
                        Cell currentCell = currentRow.getCell(rightCell, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                        currentCell.setCellStyle(AllBorderCellStyle);
                    }
                }
                // Auto-size columns at the end
                for (int rightCell = 0; rightCell < lastCell; rightCell++) {
                    doingSheet.autoSizeColumn(rightCell);
                }
            }

            int sheetNumberofHBsAg=0;
            int sheetNumberofHBsAgFinal=0;
            int sheetNumberofWidal=0;
            int sheetNumberofWidalFinal=0;
            for(int i=0;i<newSheetBook.getNumberOfSheets ();i++){
                if (newSheetBook.getSheetAt (i).getSheetName ().equals ("HBsAg")){
                    sheetNumberofHBsAg=i;
                }
                if (newSheetBook.getSheetAt (i).getSheetName ().equals ("HBsAg Final")){
                    sheetNumberofHBsAgFinal=i;
                }
                if (newSheetBook.getSheetAt (i).getSheetName ().equals ("WIDAL")){
                    sheetNumberofWidal=i;
                }
                if (newSheetBook.getSheetAt (i).getSheetName ().equals ("WIDAL Final")){
                    sheetNumberofWidalFinal=i;
                }
            }
////          HBsAg Final
//////          create place to store value
//            Set<String> hasil = new TreeSet<> ();
//            Set<String> klpUmurXHasil = new TreeSet<>();
//            Map<String, Map<String, Integer>> countMap = new HashMap<>(); // new count map
//            Sheet sheetHBsAg = newSheetBook.getSheetAt (sheetNumberofHBsAg);
//            Sheet sheetHBsAgFinal = newSheetBook.getSheetAt (sheetNumberofHBsAgFinal);
//
////          mapping value
//            for (int row = 1; row <= sheetHBsAg.getLastRowNum(); row++) {
//                String cellKelompokUmur = sheetHBsAg.getRow(row).getCell(15).getStringCellValue();      //row header
//                String cellHasilPemeriksaan = sheetHBsAg.getRow(row).getCell(11).getStringCellValue();  //column header
//
//                hasil.add(cellKelompokUmur);
//                klpUmurXHasil.add(cellHasilPemeriksaan);
//
//                // increment count in countMap
//                if (!countMap.containsKey(cellHasilPemeriksaan)) {
//                    countMap.put(cellHasilPemeriksaan, new HashMap<>());
//                }
//                Map<String, Integer> kelompokUmurCountMap = countMap.get(cellHasilPemeriksaan);
//                if (!kelompokUmurCountMap.containsKey(cellKelompokUmur)) {
//                    kelompokUmurCountMap.put(cellKelompokUmur, 1);
//                } else {
//                    kelompokUmurCountMap.put(cellKelompokUmur, kelompokUmurCountMap.get(cellKelompokUmur) + 1);
//                }
//            }
//
////          writing cell
//            sheetHBsAgFinal.createRow(0).createCell(0).setCellValue("KLP UMUR TH");
//            int rowStart = 1;
//            for (String konten : klpUmurXHasil) {
//                sheetHBsAgFinal.getRow(0).createCell(rowStart).setCellValue(konten);
//                rowStart++;
//            }
//
////          filling row
//            rowStart = 1;
//            int lastCol = klpUmurXHasil.size() + 1;
//            for (String konten : hasil) {
//                int colStart = 1;
//                sheetHBsAgFinal.createRow(rowStart).createCell(0).setCellValue(konten);
//                int total = 0;
//                for (String item : klpUmurXHasil) {
//                    if (countMap.containsKey(item) && countMap.get(item).containsKey(konten)) {
//                        int count = countMap.get(item).get(konten);
//                        sheetHBsAgFinal.getRow(rowStart).createCell(colStart++).setCellValue(count);
//                        total += count;
//                    } else {
//                        sheetHBsAgFinal.getRow(rowStart).createCell(colStart++).setCellValue(0);
//                    }
//                }
//                sheetHBsAgFinal.getRow(rowStart).createCell(lastCol).setCellValue(total); // add row total
//                rowStart++;
//            }
//
////          add grand total to last row
//            sheetHBsAgFinal.createRow(rowStart);
//            lastCell = sheetHBsAgFinal.getRow (0).getLastCellNum ();
//            sheetHBsAgFinal.getRow (0).createCell (lastCell).setCellValue ("Grand Total");
//            sheetHBsAgFinal.getRow(rowStart).createCell(0).setCellValue("Grand Total");
//            int colStart = 1;
//            for (String item : klpUmurXHasil) {
//                int total = 0;
//                for (String konten : hasil) {
//                    if (countMap.containsKey(item) && countMap.get(item).containsKey(konten)) {
//                        total += countMap.get(item).get(konten);
//                    }
//                }
//                sheetHBsAgFinal.getRow(rowStart).createCell(colStart++).setCellValue(total);
//            }
////            sheetHBsAgFinal.getRow(rowStart).createCell(lastCol).setCellValue(sheetHBsAgFinal.getLastRowNum()); // add total number of rows
//
//
////            ----------------------------------------------------------------------------------------------------------
////          HBsAg Final
////          create place to store value
//            hasil = new TreeSet<> ();
//            klpUmurXHasil = new TreeSet<>();
//            countMap = new HashMap<>(); // new count map
//            sheetHBsAg = newSheetBook.getSheetAt (sheetNumberofHBsAg);
//            sheetHBsAgFinal = newSheetBook.getSheetAt (sheetNumberofHBsAgFinal);
//
////          mapping value
//            for (int row = 1; row <= sheetHBsAg.getLastRowNum(); row++) {
//                String cellKelompokUmur = sheetHBsAg.getRow(row).getCell(12).getStringCellValue();      //row header
//                String cellHasilPemeriksaan = sheetHBsAg.getRow(row).getCell(11).getStringCellValue();  //column header
//
//                hasil.add(cellKelompokUmur);
//                klpUmurXHasil.add(cellHasilPemeriksaan);
//
//                // increment count in countMap
//                if (!countMap.containsKey(cellHasilPemeriksaan)) {
//                    countMap.put(cellHasilPemeriksaan, new HashMap<>());
//                }
//                Map<String, Integer> kelompokUmurCountMap = countMap.get(cellHasilPemeriksaan);
//                if (!kelompokUmurCountMap.containsKey(cellKelompokUmur)) {
//                    kelompokUmurCountMap.put(cellKelompokUmur, 1);
//                } else {
//                    kelompokUmurCountMap.put(cellKelompokUmur, kelompokUmurCountMap.get(cellKelompokUmur) + 1);
//                }
//            }
//
////          writing cell
//            sheetHBsAgFinal.createRow(24).createCell(0).setCellValue("KLP UMUR TH");
//            int cellStart = 1;
//            for (String konten : klpUmurXHasil) {
//                sheetHBsAgFinal.getRow(24).createCell(cellStart).setCellValue(konten);
//                cellStart++;
//            }
//
////          filling row
//            rowStart = 25;
//            lastCol = klpUmurXHasil.size() + 1;
//            for (String konten : hasil) {
//                colStart = 1;
//                sheetHBsAgFinal.createRow(rowStart).createCell(0).setCellValue(konten);
//                int total = 0;
//                for (String item : klpUmurXHasil) {
//                    if (countMap.containsKey(item) && countMap.get(item).containsKey(konten)) {
//                        int count = countMap.get(item).get(konten);
//                        sheetHBsAgFinal.getRow(rowStart).createCell(colStart++).setCellValue(count);
//                        total += count;
//                    } else {
//                        sheetHBsAgFinal.getRow(rowStart).createCell(colStart++).setCellValue(0);
//                    }
//                }
//                sheetHBsAgFinal.getRow(rowStart).createCell(lastCol).setCellValue(total); // add row total
//                rowStart++;
//            }
//
////          add grand total to last row
//            sheetHBsAgFinal.createRow(rowStart);
//            lastCell = sheetHBsAgFinal.getRow (24).getLastCellNum ();
//            sheetHBsAgFinal.getRow (24).createCell (lastCell).setCellValue ("Grand Total");
//            sheetHBsAgFinal.getRow(rowStart).createCell(0).setCellValue("Grand Total");
//            colStart = 1;
//            for (String item : klpUmurXHasil) {
//                int total = 0;
//                for (String konten : hasil) {
//                    if (countMap.containsKey(item) && countMap.get(item).containsKey(konten)) {
//                        total += countMap.get(item).get(konten);
//                    }
//                }
//                sheetHBsAgFinal.getRow(rowStart).createCell(colStart++).setCellValue(total);
//            }
////            sheetHBsAgFinal.getRow(rowStart).createCell(lastCol).setCellValue();


//          WIDAL FINAL
//          create place to store value
            Set<String> umur = new TreeSet<>();
            TreeMap<String, String> noregKlpUmurxHasil = new TreeMap<>();
            Set<String> klpUmurXHasil = new TreeSet<>();
            Map<String, Map<String, Integer>> countMap = new HashMap<>(); // new count map

            Sheet sheetWIDAL = newSheetBook.getSheetAt(sheetNumberofWidal);
            Sheet sheetWIDALFinal = newSheetBook.getSheetAt(sheetNumberofWidalFinal);

            int counter=0;
//            This line sets up a loop that will iterate through the rows of the data in sheetWIDAL. It starts from the second row (1-based index) because the first row is often used for headers.
            for (int row = 1; row <= sheetWIDAL.getLastRowNum(); row++) {
                String cellKelompokUmur = sheetWIDAL.getRow(row).getCell(15).getStringCellValue(); // row header
                String cellHasilPemeriksaan = sheetWIDAL.getRow(row).getCell(11).getStringCellValue(); // column header
                String cellnoreg = sheetWIDAL.getRow(row).getCell(0).getStringCellValue();

//            gabung noreg tanggal, map dengan hasil. sehingga setiap noreg tanggal memiliki 1 hasil.
//            next split noreg tanggal. dan map tanggal, hasil dan count.

                String noregKlpUmur = cellnoreg + "T.T" + cellKelompokUmur;
                // Check if noreg already exists
                if (noregKlpUmurxHasil.containsKey(noregKlpUmur)) {
                    // Check if the current result is "positive" and update if it is
                    if ("positive".equalsIgnoreCase(cellHasilPemeriksaan)) {
                        noregKlpUmurxHasil.put(noregKlpUmur, "positive");

                    }
                } else {
                    noregKlpUmurxHasil.put(noregKlpUmur, cellHasilPemeriksaan);
                }

                umur.add(cellKelompokUmur);
                klpUmurXHasil.add(cellHasilPemeriksaan);
            }

            for (Map.Entry<String, String> entry : noregKlpUmurxHasil.entrySet ()) {
                String[] splitValue = entry.getKey ().split("T.T");
                String klpUmur = splitValue[1];
                String hasil = entry.getValue ();


                if (!countMap.containsKey(hasil)) {
                    countMap.put(hasil, new HashMap<>());
                }

                Map<String, Integer> kelompokUmurCountMap = countMap.get(hasil);
                if (!kelompokUmurCountMap.containsKey(klpUmur)) {
                    kelompokUmurCountMap.put(klpUmur, 1);
                } else {
                    kelompokUmurCountMap.put(klpUmur, kelompokUmurCountMap.get(klpUmur) + 1);
                }
            }

////          writing cell
            sheetWIDALFinal.createRow(0).createCell(0).setCellValue("KLP UMUR TH");
            int rowStart = 1;
            for (String konten : klpUmurXHasil) {
                sheetWIDALFinal.getRow(0).createCell(rowStart).setCellValue(konten);
                rowStart++;
            }

//          filling row
            rowStart = 1;
            int lastCol = klpUmurXHasil.size() + 1;
            for (String umurs : umur) {
                int colStart = 1;
                sheetWIDALFinal.createRow(rowStart).createCell(0).setCellValue(umurs);
                int total = 0;
                for (String hasils : klpUmurXHasil) {
                    if (countMap.containsKey(hasils) && countMap.get(hasils).containsKey(umurs)) {
                        int count = countMap.get(hasils).get(umurs);
                        sheetWIDALFinal.getRow(rowStart).createCell(colStart++).setCellValue(count);
                        total += count;
                    } else {
                        sheetWIDALFinal.getRow(rowStart).createCell(colStart++).setCellValue(0);
                    }
                }
                sheetWIDALFinal.getRow(rowStart).createCell(lastCol).setCellValue(total); // add row total
                rowStart++;
            }

//          add grand total to last row
            sheetWIDALFinal.createRow(rowStart);
            lastCell = sheetWIDALFinal.getRow (0).getLastCellNum ();
            sheetWIDALFinal.getRow (0).createCell (lastCell).setCellValue ("Grand Total");
            sheetWIDALFinal.getRow(rowStart).createCell(0).setCellValue("Grand Total");
            int colStart = 1;
            for (String item : klpUmurXHasil) {
                int total = 0;
                for (String konten : umur) {
                    if (countMap.containsKey(item) && countMap.get(item).containsKey(konten)) {
                        total += countMap.get(item).get(konten);
                    }
                }
                sheetWIDALFinal.getRow(rowStart).createCell(colStart++).setCellValue(total);
            }
//            sheetHBsAgFinal.getRow(rowStart).createCell(lastCol).setCellValue(sheetHBsAgFinal.getLastRowNum()); // add total number of rows


//            ----------------------------------------------------------------------------------------------------------
//          HBsAg Final
//          create place to store value
            umur = new TreeSet<> ();
            klpUmurXHasil = new TreeSet<>();
            countMap = new HashMap<>(); // new count map
            sheetWIDAL = newSheetBook.getSheetAt (sheetNumberofWidal);
            sheetWIDALFinal = newSheetBook.getSheetAt (sheetNumberofWidalFinal);

//          mapping value
            for (int row = 1; row <= sheetWIDAL.getLastRowNum(); row++) {
                String cellKelompokUmur = sheetWIDAL.getRow(row).getCell(12).getStringCellValue();      //row header
                String cellHasilPemeriksaan = sheetWIDAL.getRow(row).getCell(11).getStringCellValue();  //column header

                umur.add(cellKelompokUmur);
                klpUmurXHasil.add(cellHasilPemeriksaan);

                // increment count in countMap
                if (!countMap.containsKey(cellHasilPemeriksaan)) {
                    countMap.put(cellHasilPemeriksaan, new HashMap<>());
                }
                Map<String, Integer> kelompokUmurCountMap = countMap.get(cellHasilPemeriksaan);
                if (!kelompokUmurCountMap.containsKey(cellKelompokUmur)) {
                    kelompokUmurCountMap.put(cellKelompokUmur, 1);
                } else {
                    kelompokUmurCountMap.put(cellKelompokUmur, kelompokUmurCountMap.get(cellKelompokUmur) + 1);
                }
            }

//          writing cell
            sheetWIDALFinal.createRow(24).createCell(0).setCellValue("KLP UMUR TH");
            int cellStart = 1;
            for (String konten : klpUmurXHasil) {
                sheetWIDALFinal.getRow(24).createCell(cellStart).setCellValue(konten);
                cellStart++;
            }

//          filling row
            rowStart = 25;
            lastCol = klpUmurXHasil.size() + 1;
            for (String konten : umur) {
                colStart = 1;
                sheetWIDALFinal.createRow(rowStart).createCell(0).setCellValue(konten);
                int total = 0;
                for (String item : klpUmurXHasil) {
                    if (countMap.containsKey(item) && countMap.get(item).containsKey(konten)) {
                        int count = countMap.get(item).get(konten);
                        sheetWIDALFinal.getRow(rowStart).createCell(colStart++).setCellValue(count);
                        total += count;
                    } else {
                        sheetWIDALFinal.getRow(rowStart).createCell(colStart++).setCellValue(0);
                    }
                }
                sheetWIDALFinal.getRow(rowStart).createCell(lastCol).setCellValue(total); // add row total
                rowStart++;
            }

//          add grand total to last row
            sheetWIDALFinal.createRow(rowStart);
            lastCell = sheetWIDALFinal.getRow (24).getLastCellNum ();
            sheetWIDALFinal.getRow (24).createCell (lastCell).setCellValue ("Grand Total");
            sheetWIDALFinal.getRow(rowStart).createCell(0).setCellValue("Grand Total");
            colStart = 1;
            for (String item : klpUmurXHasil) {
                int total = 0;
                for (String konten : umur) {
                    if (countMap.containsKey(item) && countMap.get(item).containsKey(konten)) {
                        total += countMap.get(item).get(konten);
                    }
                }
                sheetWIDALFinal.getRow(rowStart).createCell(colStart++).setCellValue(total);
            }
//            sheetHBsAgFinal.getRow(rowStart).createCell(lastCol).setCellValue();
//            todo next add PIVOT
////          create place to store value
//            Set<String> hasil = new TreeSet<> ();
//            Set<String> nickInstXTindakan = new TreeSet<>();
//            Map<String, Map<String, Integer>> countMap = new HashMap<>(); // new count map
//
////          mapping value
//            for (int row = 1; row <= noDuplicate.getLastRowNum(); row++) {
//                String cellcrByr = noDuplicate.getRow(row).getCell(8).getStringCellValue();
//                String cellNickInst = noDuplicate.getRow(row).getCell(24).getStringCellValue();
//
//                hasil.add(cellcrByr);
//                nickInstXTindakan.add(cellNickInst);
//
//                // increment count in countMap
//                if (!countMap.containsKey(cellNickInst)) {
//                    countMap.put(cellNickInst, new HashMap<>());
//                }
//                Map<String, Integer> crBayarCountMap = countMap.get(cellNickInst);
//                if (!crBayarCountMap.containsKey(cellcrByr)) {
//                    crBayarCountMap.put(cellcrByr, 1);
//                } else {
//                    crBayarCountMap.put(cellcrByr, crBayarCountMap.get(cellcrByr) + 1);
//                }
//            }
//
////          writing cell
//            instCaraBayar.createRow(0).createCell(0).setCellValue("Tanggal");
//            int rowStart = 1;
//            for (String konten : nickInstXTindakan) {
//                instCaraBayar.getRow(0).createCell(rowStart).setCellValue(konten);
//                rowStart++;
//            }
//
////          filling row
//            rowStart = 1;
//            int lastCol = nickInstXTindakan.size() + 1;
//            for (String konten : hasil) {
//                int colStart = 1;
//                instCaraBayar.createRow(rowStart).createCell(0).setCellValue(konten);
//                int total = 0;
//                for (String item : nickInstXTindakan) {
//                    if (countMap.containsKey(item) && countMap.get(item).containsKey(konten)) {
//                        int count = countMap.get(item).get(konten);
//                        instCaraBayar.getRow(rowStart).createCell(colStart++).setCellValue(count);
//                        total += count;
//                    } else {
//                        instCaraBayar.getRow(rowStart).createCell(colStart++).setCellValue(0);
//                    }
//                }
//                instCaraBayar.getRow(rowStart).createCell(lastCol).setCellValue(total); // add row total
//                rowStart++;
//            }

//          add grand total to last row
//            instCaraBayar.createRow(rowStart);
//            int lastCell = instCaraBayar.getRow (0).getLastCellNum ();
//            instCaraBayar.getRow (0).createCell (lastCell).setCellValue ("Grand Total");
//            instCaraBayar.getRow(rowStart).createCell(0).setCellValue("Grand Total");
//            int colStart = 1;
//            for (String item : nickInstXTindakan) {
//                int total = 0;
//                for (String konten : hasil) {
//                    if (countMap.containsKey(item) && countMap.get(item).containsKey(konten)) {
//                        total += countMap.get(item).get(konten);
//                    }
//                }
//                instCaraBayar.getRow(rowStart).createCell(colStart++).setCellValue(total);
//            }
//            instCaraBayar.getRow(rowStart).createCell(lastCol).setCellValue(noDuplicate.getLastRowNum()); // add total number of rows



            System.out.println ("total sheet "+newSheetBook.getNumberOfSheets ());
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
