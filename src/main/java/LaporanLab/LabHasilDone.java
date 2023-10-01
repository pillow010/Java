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
        Pattern pattern = Pattern.compile("[\\\\/:*?\"<>|.]"); // Invalid characters for sheet names
        String fileInput = "C:\\sat work\\test\\1. input\\";
        String fileOutput = "C:\\sat work\\test\\2. output\\";
        String fileNameHasilRinci = localDate + " lab hasil rinci";
        String fileNamePertindakanNew = localDate + " lab tindakan new";
        String fileNameRegister       = localDate + " lab register";
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


        // Usage:
//        File fileHasilRinci = findFile(fileInput, fileNameHasilRinci);
//        File fileRegister = findFile(fileInput, fileNameRegister);
//        File fileTindakanNew = findFile(fileInput, fileNamePertindakanNew);
//
//        if (fileHasilRinci == null || fileRegister == null || fileTindakanNew == null) {
//            return; // Handle the case where any file is missing
//        }

        try {
            File fileHasilRinci;
            File fileRegister;
            File fileTindakanNew;
            File xlsHasilRinci  = new File (fileInput + fileNameHasilRinci + ".xls");
            File xlsxHasilRinci = new File (fileInput + fileNameHasilRinci + ".xlsx");
            File xlsPertindakanNew      = new File (fileInput + fileNamePertindakanNew + ".xls");
            File xlsxPertindakanNew     = new File (fileInput + fileNamePertindakanNew + ".xlsx");
            File xlsRegister     = new File (fileInput + fileNameRegister + ".xls");
            File xlsxRegister    = new File (fileInput + fileNameRegister + ".xlsx");

            if (xlsxHasilRinci.exists()) {
                fileHasilRinci = xlsxHasilRinci;
            } else if (xlsHasilRinci.exists()) {
                fileHasilRinci = xlsHasilRinci;
            } else {
                System.out.println("File not found: " + fileInput + fileNameHasilRinci);
                return;
            }

            if (xlsxPertindakanNew.exists()) {
                fileTindakanNew = xlsxPertindakanNew;
            } else if (xlsPertindakanNew.exists()) {
                fileTindakanNew = xlsPertindakanNew;
            } else {
                System.out.println("File not found: " + fileInput + fileNamePertindakanNew);
                return;
            }

            if (xlsRegister.exists()) {
                fileRegister = xlsRegister;
            } else if (xlsxRegister.exists()) {
                fileRegister = xlsxRegister;
            } else {
                System.out.println("File not found: " + fileInput + fileNameRegister);
                return;
            }

            bookHasilRinci = WorkbookFactory.create (new FileInputStream (fileHasilRinci));
            Workbook bookRegister = WorkbookFactory.create (new FileInputStream (fileRegister));
            Workbook bookPertindakan = WorkbookFactory.create (new FileInputStream (fileTindakanNew));
//            InputStream hasilRinci = new FileInputStream(fileHasilRinci);
//            bookHasilRinci = WorkbookFactory.create (hasilRinci);
//            InputStream register = new FileInputStream(fileRegister);
//            Workbook bookRegister = WorkbookFactory.create (register);
//            FileInputStream inputStream = new FileInputStream(fileTindakanNew);
//            Workbook bookPertindakan = WorkbookFactory.create(inputStream);
            Sheet sheetRegister = bookRegister.getSheetAt (0);
            Sheet sheetPerTindakan = bookPertindakan.getSheetAt (0);
            Sheet sheetHasilRinci = bookHasilRinci.getSheetAt(0);
            int lastCell = sheetHasilRinci.getRow(0).getLastCellNum();
            sheetHasilRinci.getRow (0).createCell (lastCell).setCellValue ("Diagnosa");
            sheetHasilRinci.getRow (0).createCell (lastCell+1).setCellValue ("NIK");

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

            // Create a HashMap to store the keys and corresponding values from sheetPerTindakan
            HashMap<String, String> registerMap = new HashMap<> ();
            for (int j = 1; j <= sheetRegister.getLastRowNum(); j++) {
                String key = sheetRegister.getRow (j).getCell (0).getStringCellValue () +
                        sheetRegister.getRow (j).getCell (1).getStringCellValue () +
                        sheetRegister.getRow (j).getCell (2).getStringCellValue () +
                        sheetRegister.getRow (j).getCell (3).getStringCellValue () +
                        sheetRegister.getRow (j).getCell (4).getStringCellValue ();
                Cell cellValue = sheetRegister.getRow (j).getCell (11);
                if (cellValue == null) {
                    registerMap.put (key, "");
                } else {
                    String value = cellValue.getStringCellValue ();
                    registerMap.put (key, value);
                }
            }

            // Iterate through sheetHasilRinci and perform lookups in the HashMap
            for (int i = 1; i <= sheetHasilRinci.getLastRowNum(); i++) {
                String noreg = sheetHasilRinci.getRow(i).getCell(0).getStringCellValue().replaceAll(pattern.pattern(), "");
                sheetHasilRinci.getRow(i).getCell(0).setCellValue(noreg);
                Cell cellHasil = sheetHasilRinci.getRow (i).getCell (11);
                Cell cellPemeriksaan = sheetHasilRinci.getRow (i).getCell (9);
                if (cellHasil.getStringCellValue ().contains ("/") && cellPemeriksaan.getStringCellValue ().equalsIgnoreCase ("widal")){
                    cellHasil.setCellValue ("positive");
                }
                String keyToLookup = sheetHasilRinci.getRow(i).getCell(0).getStringCellValue();
                String valueFromMap = perTindakanMap.get(keyToLookup);
                String nikFromMap = registerMap.get (keyToLookup);

                if (valueFromMap != null) {
                    sheetHasilRinci.getRow(i).createCell(lastCell).setCellValue(valueFromMap);
                    sheetHasilRinci.getRow (i).createCell (lastCell+1).setCellValue (nikFromMap);
                }
            }

            int lastRow = sheetHasilRinci.getLastRowNum();
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
//          HBsAg Final
//          create place to store value
            Set<String> hasil = new TreeSet<> ();
            Set<String> klpUmurXHasil = new TreeSet<>();
            Map<String, Map<String, Integer>> countMap = new HashMap<>(); // new count map
            Sheet sheetHBsAg = newSheetBook.getSheetAt (sheetNumberofHBsAg);
            Sheet sheetHBsAgFinal = newSheetBook.getSheetAt (sheetNumberofHBsAgFinal);

//          mapping value
            for (int row = 1; row <= sheetHBsAg.getLastRowNum(); row++) {
                String cellKelompokUmur = sheetHBsAg.getRow(row).getCell(15).getStringCellValue();      //row header
                String cellHasilPemeriksaan = sheetHBsAg.getRow(row).getCell(11).getStringCellValue();  //column header

                hasil.add(cellKelompokUmur);
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
            sheetHBsAgFinal.createRow(0).createCell(0).setCellValue("KLP UMUR TH");
            int rowStart = 1;
            for (String konten : klpUmurXHasil) {
                sheetHBsAgFinal.getRow(0).createCell(rowStart).setCellValue(konten);
                rowStart++;
            }

//          filling row
            rowStart = 1;
            int lastCol = klpUmurXHasil.size() + 1;
            for (String konten : hasil) {
                int colStart = 1;
                sheetHBsAgFinal.createRow(rowStart).createCell(0).setCellValue(konten);
                int total = 0;
                for (String item : klpUmurXHasil) {
                    if (countMap.containsKey(item) && countMap.get(item).containsKey(konten)) {
                        int count = countMap.get(item).get(konten);
                        sheetHBsAgFinal.getRow(rowStart).createCell(colStart++).setCellValue(count);
                        total += count;
                    } else {
                        sheetHBsAgFinal.getRow(rowStart).createCell(colStart++).setCellValue(0);
                    }
                }
                sheetHBsAgFinal.getRow(rowStart).createCell(lastCol).setCellValue(total); // add row total
                rowStart++;
            }

//          add grand total to last row
            sheetHBsAgFinal.createRow(rowStart);
            lastCell = sheetHBsAgFinal.getRow (0).getLastCellNum ();
            sheetHBsAgFinal.getRow (0).createCell (lastCell).setCellValue ("Grand Total");
            sheetHBsAgFinal.getRow(rowStart).createCell(0).setCellValue("Grand Total");
            int colStart = 1;
            for (String item : klpUmurXHasil) {
                int total = 0;
                for (String konten : hasil) {
                    if (countMap.containsKey(item) && countMap.get(item).containsKey(konten)) {
                        total += countMap.get(item).get(konten);
                    }
                }
                sheetHBsAgFinal.getRow(rowStart).createCell(colStart++).setCellValue(total);
            }


//            ----------------------------------------------------------------------------------------------------------
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

            // Initialize data structures
            Set<String> hasilSet = new TreeSet<>();
            Set<String> klpUmurXHasilSet = new TreeSet<>();
            countMap = new HashMap<>();

            // Process the data from the sheet
            for (int row = 1; row <= sheetHBsAg.getLastRowNum(); row++) {
                String cellKelompokUmur = sheetHBsAg.getRow(row).getCell(12).getStringCellValue();
                String cellHasilPemeriksaan = sheetHBsAg.getRow(row).getCell(11).getStringCellValue();

                hasilSet.add(cellKelompokUmur);
                klpUmurXHasilSet.add(cellHasilPemeriksaan);

                countMap.computeIfAbsent(cellHasilPemeriksaan, k -> new HashMap<>());
                Map<String, Integer> kelompokUmurCountMap = countMap.get(cellHasilPemeriksaan);

                kelompokUmurCountMap.merge(cellKelompokUmur, 1, Integer::sum);
            }

            // Write cell data
            rowStart = 24;
            colStart = 1;
            lastCol = klpUmurXHasilSet.size() + 1;

            sheetHBsAgFinal.createRow(rowStart).createCell(0).setCellValue("KLP UMUR TH");
            for (String konten : klpUmurXHasilSet) {
                sheetHBsAgFinal.getRow(rowStart).createCell(colStart).setCellValue(konten);
                colStart++;
            }

            // Fill rows with counts
            rowStart = 25;
            for (String stringHasil : hasilSet) {
                colStart = 1;
                sheetHBsAgFinal.createRow(rowStart).createCell(0).setCellValue(stringHasil);
                int total = 0;
                for (String item : klpUmurXHasilSet) {
                    total += countMap.getOrDefault(item, Collections.emptyMap()).getOrDefault(stringHasil, 0);
                    sheetHBsAgFinal.getRow(rowStart).createCell(colStart).setCellValue(total);
                    colStart++;
                }
                sheetHBsAgFinal.getRow(rowStart).createCell(lastCol).setCellValue(total); // add row total
                rowStart++;
            }

            // Add grand total to last row
            sheetHBsAgFinal.createRow(rowStart);
            colStart = 1;
            for (String item : klpUmurXHasilSet) {
                Map<String, Map<String, Integer>> finalCountMap2 = countMap;
                int total = hasilSet.stream().mapToInt(stringHasil -> finalCountMap2.getOrDefault(item, Collections.emptyMap()).getOrDefault(stringHasil, 0)).sum();
                sheetHBsAgFinal.getRow(rowStart).createCell(colStart).setCellValue(total);
                colStart++;
            }

            // Calculate and set the grand total
            Map<String, Map<String, Integer>> finalCountMap3 = countMap;
            int grandTotal = klpUmurXHasilSet.stream()
                    .mapToInt(item -> hasilSet.stream().mapToInt(stringHasil -> finalCountMap3.getOrDefault(item, Collections.emptyMap()).getOrDefault(stringHasil, 0)).sum())
                    .sum();

            sheetHBsAgFinal.getRow(rowStart).createCell(0).setCellValue("Grand Total");
            sheetHBsAgFinal.getRow(rowStart).createCell(lastCol).setCellValue(grandTotal);



//          WIDAL FINAL
//          create place to store value
            Set<String> umur = new TreeSet<>();
            TreeMap<String, String> noregKlpUmurxHasil = new TreeMap<>();
            klpUmurXHasil = new TreeSet<>();
            countMap = new HashMap<>(); // new count map

            Sheet sheetWIDAL = newSheetBook.getSheetAt(sheetNumberofWidal);
            Sheet sheetWIDALFinal = newSheetBook.getSheetAt(sheetNumberofWidalFinal);



            // Iterate through the data
            for (int row = 1; row <= sheetWIDAL.getLastRowNum(); row++) {
                String cellKelompokUmur = sheetWIDAL.getRow(row).getCell(15).getStringCellValue(); // Extract age group
                String cellHasilPemeriksaan = sheetWIDAL.getRow(row).getCell(11).getStringCellValue(); // Extract test result
                String cellnoreg = sheetWIDAL.getRow(row).getCell(0).getStringCellValue(); // Extract registration number

                // Combine registration number and age group into a single key
                String noregKlpUmur = cellnoreg + "T.T" + cellKelompokUmur;

                // Check if the combined key already exists in noregKlpUmurxHasil
                if (noregKlpUmurxHasil.containsKey(noregKlpUmur)) {
                    // Check if the current result is "positive" and update if it is
                    if (cellHasilPemeriksaan.equalsIgnoreCase ("positive")) {
                        noregKlpUmurxHasil.put (noregKlpUmur, "positive");
                    }
                } else {
                    noregKlpUmurxHasil.put(noregKlpUmur, cellHasilPemeriksaan);
                }

                // Add unique age groups and test results to respective sets
                umur.add(cellKelompokUmur);
                klpUmurXHasil.add(cellHasilPemeriksaan);
            }

            // Calculate counts for each test result and age group
            for (Map.Entry<String, String> entry : noregKlpUmurxHasil.entrySet()) {
                String[] splitValue = entry.getKey().split("T.T");
                String klpUmur = splitValue[1];
                String hasil1 = entry.getValue();

                // Ensure that the result exists in countMap
                if (!countMap.containsKey(hasil1)) {
                    countMap.put(hasil1, new HashMap<>());
                }

                Map<String, Integer> kelompokUmurCountMap = countMap.get(hasil1);

                // Increment the count for the current age group
                if (!kelompokUmurCountMap.containsKey(klpUmur)) {
                    kelompokUmurCountMap.put(klpUmur, 1);
                } else {
                    kelompokUmurCountMap.put(klpUmur, kelompokUmurCountMap.get(klpUmur) + 1);
                }
            }

            // Write cell data
            sheetWIDALFinal.createRow(0).createCell(0).setCellValue("KLP UMUR TH");
            rowStart = 1;
            for (String konten : klpUmurXHasil) {
                sheetWIDALFinal.getRow(0).createCell(rowStart).setCellValue(konten);
                rowStart++;
            }

            // Fill rows with counts
            rowStart = 1;
            lastCol = klpUmurXHasil.size() + 1;
            for (String umurs : umur) {
                colStart = 1;
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
                sheetWIDALFinal.getRow(rowStart).createCell(lastCol).setCellValue(total); // Add row total
                rowStart++;
            }

            // Add grand total to last row
            sheetWIDALFinal.createRow(rowStart);
            lastCell = sheetWIDALFinal.getRow(0).getLastCellNum();
            sheetWIDALFinal.getRow(0).createCell(lastCell).setCellValue("Grand Total");
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
            sheetWIDALFinal.getRow(rowStart).createCell(0).setCellValue("Grand Total");


//            ----------------------------------------------------------------------------------------------------------
////          HBsAg Final
////          create place to store value
//            umur = new TreeSet<> ();
//            klpUmurXHasil = new TreeSet<>();
//            countMap = new HashMap<>(); // new count map
//            sheetWIDAL = newSheetBook.getSheetAt (sheetNumberofWidal);
//            sheetWIDALFinal = newSheetBook.getSheetAt (sheetNumberofWidalFinal);
//
////          mapping value
//            for (int row = 1; row <= sheetWIDAL.getLastRowNum(); row++) {
//                String cellKelompokUmur = sheetWIDAL.getRow(row).getCell(12).getStringCellValue();      //row header
//                String cellHasilPemeriksaan = sheetWIDAL.getRow(row).getCell(11).getStringCellValue();  //column header
//                String cellnoreg = sheetWIDAL.getRow(row).getCell(0).getStringCellValue(); // Extract registration number
//
//                // Combine registration number and age group into a single key
//                String noregKlpUmur = cellnoreg + "T.T" + cellKelompokUmur;
//
//                // Check if the combined key already exists in noregKlpUmurxHasil
//                if (noregKlpUmurxHasil.containsKey(noregKlpUmur)) {
//                    // Check if the current result is "positive" and update if it is
//                    if (cellHasilPemeriksaan.equalsIgnoreCase ("positive")) {
//                        noregKlpUmurxHasil.put (noregKlpUmur, "positive");
//                    }
//                } else {
//                    noregKlpUmurxHasil.put(noregKlpUmur, cellHasilPemeriksaan);
//                }
//
//                // Add unique age groups and test results to respective sets
//                umur.add(cellKelompokUmur);
//                klpUmurXHasil.add(cellHasilPemeriksaan);
//            }
//
//            // Calculate counts for each test result and age group
//            for (Map.Entry<String, String> entry : noregKlpUmurxHasil.entrySet()) {
//                String[] splitValue = entry.getKey().split("T.T");
//                String klpUmur = splitValue[1];
//                String hasil1 = entry.getValue();
//
//                // Ensure that the result exists in countMap
//                if (!countMap.containsKey(hasil1)) {
//                    countMap.put(hasil1, new HashMap<>());
//                }
//
//                Map<String, Integer> kelompokUmurCountMap = countMap.get(hasil1);
//
//                // Increment the count for the current age group
//                if (!kelompokUmurCountMap.containsKey(klpUmur)) {
//                    kelompokUmurCountMap.put(klpUmur, 1);
//                } else {
//                    kelompokUmurCountMap.put(klpUmur, kelompokUmurCountMap.get(klpUmur) + 1);
//                }
//            }
//
//            // Write cell data
//            sheetWIDALFinal.createRow(24).createCell(0).setCellValue("KLP UMUR TH");
//            rowStart = 1;
//            for (String konten : klpUmurXHasil) {
//                sheetWIDALFinal.getRow(24).createCell(rowStart).setCellValue(konten);
//                rowStart++;
//            }
//
//            // Fill rows with counts
//            rowStart = 25;
//            lastCol = klpUmurXHasil.size() + 1;
//            for (String umurs : umur) {
//                colStart = 1;
//                sheetWIDALFinal.createRow(rowStart).createCell(0).setCellValue(umurs);
//                int total = 0;
//                for (String hasils : klpUmurXHasil) {
//                    if (countMap.containsKey(hasils) && countMap.get(hasils).containsKey(umurs)) {
//                        int count = countMap.get(hasils).get(umurs);
//                        sheetWIDALFinal.getRow(rowStart).createCell(colStart++).setCellValue(count);
//                        total += count;
//                    } else {
//                        sheetWIDALFinal.getRow(rowStart).createCell(colStart++).setCellValue(0);
//                    }
//                }
//                sheetWIDALFinal.getRow(rowStart).createCell(lastCol).setCellValue(total); // Add row total
//                rowStart++;
//            }
//
//            // Add grand total to last row
//            sheetWIDALFinal.createRow(rowStart);
//            lastCell = sheetWIDALFinal.getRow(24).getLastCellNum();
//            sheetWIDALFinal.getRow(24).createCell(lastCell).setCellValue("Grand Total");
//            sheetWIDALFinal.getRow(rowStart).createCell(0).setCellValue("Grand Total");
//            colStart = 1;
//            for (String item : klpUmurXHasil) {
//                int total = 0;
//                for (String konten : umur) {
//                    if (countMap.containsKey(item) && countMap.get(item).containsKey(konten)) {
//                        total += countMap.get(item).get(konten);
//                    }
//                }
//                sheetWIDALFinal.getRow(rowStart).createCell(colStart++).setCellValue(total);
//            }
//            int grandtotal=0;
//            for (int i=1;i<sheetWIDALFinal.getRow (rowStart).getLastCellNum ();i++){
//                int value = (int)(sheetWIDALFinal.getRow(rowStart).getCell(i).getNumericCellValue ());
//                grandtotal += value;
//            }
//            sheetWIDALFinal.getRow (rowStart).createCell (klpUmurXHasil.size ()+1).setCellValue (grandtotal);

            // Initialize data structures
            Set<String> umurSet = new TreeSet<>();
            klpUmurXHasilSet = new TreeSet<>();
            countMap = new HashMap<>();

            // Process the data from the sheet
            for (int row = 1; row <= sheetWIDAL.getLastRowNum(); row++) {
                String cellKelompokUmur = sheetWIDAL.getRow(row).getCell(12).getStringCellValue();
                String cellHasilPemeriksaan = sheetWIDAL.getRow(row).getCell(11).getStringCellValue();
                String cellNoReg = sheetWIDAL.getRow(row).getCell(0).getStringCellValue();

                String noregKlpUmur = cellNoReg + "T.T" + cellKelompokUmur;

                if (noregKlpUmurxHasil.containsKey(noregKlpUmur)) {
                    if (cellHasilPemeriksaan.equalsIgnoreCase("positive")) {
                        noregKlpUmurxHasil.put(noregKlpUmur, "positive");
                    }
                } else {
                    noregKlpUmurxHasil.put(noregKlpUmur, cellHasilPemeriksaan);
                }

                umurSet.add(cellKelompokUmur);
                klpUmurXHasilSet.add(cellHasilPemeriksaan);
            }

            // Calculate counts for each test result and age group
            for (Map.Entry<String, String> entry : noregKlpUmurxHasil.entrySet()) {
                String[] splitValue = entry.getKey().split("T.T");
                String klpUmur = splitValue[1];
                String stringHasil = entry.getValue();

                countMap.computeIfAbsent(stringHasil, k -> new HashMap<>());
                countMap.get(stringHasil).merge(klpUmur, 1, Integer::sum);
            }

            // Write cell data
            rowStart = 24;
            colStart = 1;
            lastCol = klpUmurXHasilSet.size() + 1;

            sheetWIDALFinal.createRow(rowStart).createCell(0).setCellValue("KLP UMUR TH");
            for (String konten : klpUmurXHasilSet) {
                sheetWIDALFinal.getRow(rowStart).createCell(colStart).setCellValue(konten);
                colStart++;
            }

            // Fill rows with counts
            rowStart = 25;
            for (String stringUmur : umurSet) {
                colStart = 1;
                sheetWIDALFinal.createRow(rowStart).createCell(0).setCellValue(stringUmur);
                int total = 0;
                for (String stringHasil : klpUmurXHasilSet) {
                    total += countMap.getOrDefault(stringHasil, Collections.emptyMap()).getOrDefault(stringUmur, 0);
                    sheetWIDALFinal.getRow(rowStart).createCell(colStart).setCellValue(total);
                    colStart++;
                }
                sheetWIDALFinal.getRow(rowStart).createCell(lastCol).setCellValue(total);
                rowStart++;
            }

            // Add grand total to last row
            sheetWIDALFinal.createRow(rowStart);
            colStart = 1;
            for (String item : klpUmurXHasilSet) {
                Map<String, Map<String, Integer>> finalCountMap = countMap;
                int total = umurSet.stream().mapToInt(intUmur -> finalCountMap.getOrDefault(item, Collections.emptyMap()).getOrDefault(intUmur, 0)).sum();
                sheetWIDALFinal.getRow(rowStart).createCell(colStart).setCellValue(total);
                colStart++;
            }

            Map<String, Map<String, Integer>> finalCountMap1 = countMap;
            grandTotal = klpUmurXHasilSet.stream()
                    .mapToInt(item -> umurSet.stream().mapToInt(intUmur -> finalCountMap1.getOrDefault(item, Collections.emptyMap()).getOrDefault(intUmur, 0)).sum())
                    .sum();

            sheetWIDALFinal.getRow(rowStart).createCell(0).setCellValue("Grand Total");
            sheetWIDALFinal.getRow(rowStart).createCell(lastCol).setCellValue(grandTotal);


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
    private static File findFile(String fileInput, String fileName) {
        File xlsxFile = new File(fileInput + fileName + ".xlsx");
        File xlsFile = new File(fileInput + fileName + ".xls");

        if (xlsxFile.exists()) {
            return xlsxFile;
        } else if (xlsFile.exists()) {
            return xlsFile;
        } else {
            System.out.println("File not found: " + fileInput + fileName);
            return null; // or throw an exception if you prefer
        }
    }

}
