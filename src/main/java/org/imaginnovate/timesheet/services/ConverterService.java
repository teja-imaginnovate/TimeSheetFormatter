package org.imaginnovate.timesheet.services;

import com.joestelmach.natty.DateGroup;
import com.joestelmach.natty.Parser;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.*;


public class ConverterService {

    private static final String RAW_DATE_FORMAT = "MMM dd, yyyy";
    private static final String OUT_PUT_DATE_FORMAT = "dd-MMM-yyyy";
    private static final String CALI_BRI_STYLE = "Calibri";


    public byte[] generateFromCSV(Workbook workbook) throws IOException {
        return generate(workbook);
    }

    private byte[] generate(Workbook inputSheet) {
        try {

            Sheet sourceSheet = inputSheet.getSheetAt(0);
            LocalDate date = this.parseDateFromCell(sourceSheet.getRow(1).getCell(3));
            int month = date.getMonth().getValue();
            int year = date.getYear();
            int rowCount = sourceSheet.getPhysicalNumberOfRows();
            System.out.println("Total rows found - " + rowCount);

            // Changing the date format to match with report format
            for (int rowIndex = 1; rowIndex < rowCount; rowIndex++) {
                Cell dateCell = this.findDateFromSheet(sourceSheet, rowIndex);
                String inputDate = dateCell.toString();
                LocalDate localDate = parseDateFromCellValue(inputDate);
                DateTimeFormatter outputFormatter = DateTimeFormatter.ofPattern(OUT_PUT_DATE_FORMAT);
                String formattedDateStr = localDate.format(outputFormatter);
                dateCell.setCellValue(formattedDateStr);
            }

            TreeMap<String, Double> employeeNames = this.findAllEmployeeNames(sourceSheet);
            Workbook outputWorkbook = new XSSFWorkbook();
            CellStyle borderStyle = outputWorkbook.createCellStyle();
            Font boldFont = outputWorkbook.createFont();
            boldFont.setFontName(CALI_BRI_STYLE);
            borderStyle.setBorderTop(BorderStyle.THIN);
            borderStyle.setBorderBottom(BorderStyle.THIN);
            borderStyle.setBorderLeft(BorderStyle.THIN);
            borderStyle.setBorderRight(BorderStyle.THIN);
            borderStyle.setFont(boldFont);
            borderStyle.setAlignment(HorizontalAlignment.LEFT);
            borderStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            borderStyle.setWrapText(true);
            this.addSummaryPage(outputWorkbook, employeeNames, borderStyle);
            this.addEachTimeSheet(outputWorkbook, employeeNames, sourceSheet, month, year, borderStyle);
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            outputWorkbook.write(outputStream);
            return outputStream.toByteArray();
        } catch (Exception var15) {
            System.out.println(var15.getMessage());
            return null;
        }
    }


    private LocalDate parseDateFromCell(Cell cell) {
        if (cell == null) {
            return null;
        } else {
            return getLocalDate(cell.getStringCellValue());
        }
    }

    private LocalDate parseDateFromCellValue(String cellValue) {
        if (cellValue == null) {
            return null;
        } else {
            return getLocalDate(cellValue);
        }
    }

    private static LocalDate getLocalDate(String input) {
        Parser parser = new Parser();
        List<DateGroup> groups = parser.parse(input);
        Date date = groups.get(0).getDates().get(0);
        return date.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
    }

    private void addSummaryPage(Workbook outputWorkbook, Map<String, Double> employeeNames, CellStyle borderStyle) {
        Sheet outputSheet = outputWorkbook.createSheet("Summary");
        String[] summaryColumns = new String[]{"Names", "Hours", "New/Existing"};
        CellStyle totalStyle = outputWorkbook.createCellStyle();
        totalStyle.setAlignment(HorizontalAlignment.LEFT);
        Font boldFont = outputWorkbook.createFont();
        boldFont.setFontName(CALI_BRI_STYLE);
        boldFont.setBold(true);
        totalStyle.setFont(boldFont);
        this.addColumns(summaryColumns, outputSheet);
        this.fillSummarySheet(outputSheet, employeeNames);
        this.fitColumnContent(summaryColumns.length, outputSheet);
        this.addBorders(outputSheet, borderStyle, summaryColumns.length);
        byte[] googleBlueRGB = new byte[]{66, -123, -12};
        XSSFColor xssfColor = new XSSFColor(googleBlueRGB, null);
        this.applyColourFontOnColumns(outputWorkbook, outputSheet, summaryColumns.length, xssfColor);
        int lastRowNum = outputSheet.getLastRowNum();
        Row lastRow = outputSheet.getRow(lastRowNum);
        if (lastRow != null) {
            for (int col = 0; col < summaryColumns.length; ++col) {
                Cell cell = lastRow.getCell(col);
                if (cell == null) {
                    cell = lastRow.createCell(col);
                }

                cell.setCellStyle(totalStyle);
            }
        }

    }

    private void addEachTimeSheet(Workbook outputWorkbook, Map<String, Double> employeeNames, Sheet sourceSheet, int month, int year, CellStyle style) {
        String[] columns = new String[]{"Name", "Date", "Title", "Description", "Project Time"};

        for (String name : employeeNames.keySet()) {
            Sheet currentSheet = outputWorkbook.createSheet(name);
            this.addColumns(columns, currentSheet);
            currentSheet.setColumnWidth(3, 12800);
            this.addEachPersonSheetData(style, sourceSheet, currentSheet, name, month, year);
            this.fitColumnContent(columns.length, currentSheet);
            this.addBorders(currentSheet, style, columns.length);
            byte[] googleBlueRGB = new byte[]{66, -123, -12};
            XSSFColor xssfColor = new XSSFColor(googleBlueRGB, null);
            this.applyColourFontOnColumns(outputWorkbook, currentSheet, columns.length, xssfColor);
            this.updateWeekendColour(outputWorkbook, currentSheet, columns.length);
            int lastRowNum = currentSheet.getLastRowNum();
            currentSheet.createRow(lastRowNum + 1);
            double totalHours = 0.0;

            for (int i = 1; i <= lastRowNum; ++i) {
                Row row = currentSheet.getRow(i);
                if (row != null) {
                    Cell hoursCell = row.getCell(4);
                    if (hoursCell != null) {
                        String value = hoursCell.toString();

                        try {
                            totalHours += value.isEmpty() ? 0.0 : Double.parseDouble(value);
                        } catch (NumberFormatException var23) {
                            System.out.println(var23.getMessage());
                        }
                    }
                }
            }

            CellStyle totalStyle = outputWorkbook.createCellStyle();
            totalStyle.setAlignment(HorizontalAlignment.LEFT);
            Font boldFont = outputWorkbook.createFont();
            boldFont.setBold(true);
            totalStyle.setFont(boldFont);
            Row totalRow = currentSheet.createRow(lastRowNum + 2);
            Cell labelCell = totalRow.createCell(0);
            labelCell.setCellValue("Total");
            Cell totalCell = totalRow.createCell(4);
            totalCell.setCellValue(totalHours);

            for (int col = 0; col < columns.length; ++col) {
                Cell cell = totalRow.getCell(col);
                if (cell == null) {
                    cell = totalRow.createCell(col);
                }

                cell.setCellStyle(totalStyle);
            }
        }

    }

    private void addBorders(Sheet sheet, CellStyle style, int totalColumns) {

        for (Row row : sheet) {
            for (int column = 0; column < totalColumns; ++column) {
                if (row.getCell(column) != null) {
                    row.getCell(column).setCellStyle(style);
                }
            }
        }

    }

    private void updateWeekendColour(Workbook workbook, Sheet sheet, int length) {
        int rowIndex = 0;
        sheet.setColumnWidth(0, 6400);
        sheet.setColumnWidth(1, 4352);
        sheet.setColumnWidth(2, 5120);
        sheet.setColumnWidth(3, 30720);

        for (Row row : sheet) {
            if (rowIndex == 0) {
                ++rowIndex;
            } else {
                Cell cell = row.getCell(1);
                LocalDate date = null;
                String cellValue = cell.toString();

                try {
                    date = parseDateFromCellValue(cellValue);
                } catch (DateTimeParseException var16) {
                    System.out.println(var16.getMessage());
                }

                CellStyle descStyle = workbook.createCellStyle();
                Font boldFont = workbook.createFont();
                boldFont.setFontName(CALI_BRI_STYLE);
                descStyle.setBorderTop(BorderStyle.THIN);
                descStyle.setBorderBottom(BorderStyle.THIN);
                descStyle.setBorderLeft(BorderStyle.THIN);
                descStyle.setBorderRight(BorderStyle.THIN);
                descStyle.setFont(boldFont);
                descStyle.setAlignment(HorizontalAlignment.LEFT);
                descStyle.setVerticalAlignment(VerticalAlignment.CENTER);
                Cell descriptionCell = row.getCell(3);
                if (row.getCell(3).equals(descriptionCell)) {
                    row.setHeight((short) -1);
                }

                if (descriptionCell != null) {
                    descriptionCell.setCellStyle(descStyle);
                }

                byte[] cyanRGB;
                XSSFColor xssfColor;
                assert date != null;
                if (this.isWeekend(date)) {
                    cyanRGB = new byte[]{-109, -60, 125};
                    xssfColor = new XSSFColor(cyanRGB, null);
                    this.applyColour(workbook, sheet, length, rowIndex, xssfColor);
                } else if (descriptionCell.toString().isEmpty()) {
                    descriptionCell.setCellValue("On Leave");
                    cyanRGB = new byte[]{0, -1, -1};
                    xssfColor = new XSSFColor(cyanRGB, null);
                    this.applyColour(workbook, sheet, length, rowIndex, xssfColor);
                }

                ++rowIndex;
            }
        }

    }

    private void applyColourFontOnColumns(Workbook outputWorkbook, Sheet sheet, int length, XSSFColor xssfColor) {
        Font boldFont = outputWorkbook.createFont();
        boldFont.setFontName(CALI_BRI_STYLE);
        boldFont.setBold(true);
        XSSFCellStyle headerStyle = ((XSSFWorkbook) outputWorkbook).createCellStyle();
        headerStyle.setBorderTop(BorderStyle.THIN);
        headerStyle.setBorderBottom(BorderStyle.THIN);
        headerStyle.setBorderLeft(BorderStyle.THIN);
        headerStyle.setBorderRight(BorderStyle.THIN);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerStyle.setFont(boldFont);
        headerStyle.setWrapText(true);
        headerStyle.setFillForegroundColor(xssfColor);
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Row firstRow = sheet.getRow(0);

        for (int column = 0; column < length; ++column) {
            Cell cell = firstRow.getCell(column);
            cell.setCellStyle(headerStyle);
        }

    }

    private void applyColour(Workbook outputWorkbook, Sheet sheet, int length, int rowIndex, XSSFColor xssfColor) {
        XSSFRow row = (XSSFRow) sheet.getRow(rowIndex);
        if (row != null) {
            for (int column = 0; column < length; ++column) {
                Cell cell = row.getCell(column);
                if (cell != null) {
                    CellStyle originalStyle = cell.getCellStyle();
                    XSSFCellStyle newStyle = ((XSSFWorkbook) outputWorkbook).createCellStyle();
                    if (originalStyle != null) {
                        newStyle.cloneStyleFrom(originalStyle);
                    }

                    newStyle.setFillForegroundColor(xssfColor);
                    newStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    cell.setCellStyle(newStyle);
                }
            }

        }
    }

    private void fitColumnContent(int length, Sheet sheet) {
        for (int column = 0; column < length; ++column) {
            sheet.autoSizeColumn(column);
        }

    }

    private void addEachPersonSheetData(CellStyle style, Sheet sourceSheet, Sheet destinationSheet, String name, int month, int year) {
        LocalDate firstDate = LocalDate.of(year, month, 1);
        LocalDate lastDate = firstDate.withDayOfMonth(firstDate.lengthOfMonth());
        int rowIndex = 1;

        for (LocalDate date = firstDate; !date.isAfter(lastDate); date = date.plusDays(1L)) {
            Row row = destinationSheet.createRow(rowIndex++);
            Cell nameCell = row.createCell(0);
            nameCell.setCellValue(name);
            nameCell.setCellStyle(style);
            Cell dateCell = row.createCell(1);
            DateTimeFormatter formatter = DateTimeFormatter.ofPattern(OUT_PUT_DATE_FORMAT);
            dateCell.setCellValue(date.format(formatter));
            Cell titleCell = row.createCell(2);
            titleCell.setCellStyle(style);
            Cell descriptionCell = row.createCell(3);
            descriptionCell.setCellStyle(style);
            row.createCell(4);
            Set<String> titles = new HashSet<>();
            if (!this.isWeekend(date)) {
                for (Row sourceRow : sourceSheet) {
                    if (sourceRow.getRowNum() != 0) {
                        Cell empCell = sourceRow.getCell(1);
                        Cell dateCellSrc = sourceRow.getCell(3);
                        if (empCell != null && dateCellSrc != null && name.equals(empCell.toString())) {
                            try {
                                LocalDate srcDate = parseDateFromCellValue(dateCellSrc.toString());
                                if (srcDate.equals(date)) {
                                    Cell titleDataCell = sourceRow.getCell(4);
                                    if (titleDataCell != null) {
                                        String titleStr = titleDataCell.getStringCellValue();
                                        if (titleStr.startsWith("'")) {
                                            titleStr = titleStr.substring(1);
                                        }
                                        if (!titleStr.isEmpty()) {
                                            titles.add(titleStr);
                                        }
                                    }
                                }
                            } catch (Exception var31) {
                                System.out.println(var31.getMessage());
                            }
                        }
                    }
                }
                String concatenatedTitles = String.join(", ", titles);
                titleCell.setCellValue(concatenatedTitles);
            }
        }

        this.updateDescriptionAndHours(sourceSheet, destinationSheet, name, style);
    }

    private int findColumnIndex(Sheet sourceSheet, String fieldName) {
        Row row = sourceSheet.getRow(0);
        int columnIndex = 0;

        for (Iterator<Cell> var5 = row.iterator(); var5.hasNext(); ++columnIndex) {
            Cell cell = var5.next();
            if (cell.toString().equals(fieldName)) {
                return columnIndex;
            }
        }

        return -1;
    }

    private String findNameFromSheet(Sheet sourceSheet, int rowIndex) {
        int colIndex = this.findColumnIndex(sourceSheet, "Emp Name");
        return sourceSheet.getRow(rowIndex).getCell(colIndex).toString();
    }

    private Cell findDateFromSheet(Sheet sourceSheet, int rowIndex) {
        int colIndex = this.findColumnIndex(sourceSheet, "Date");
        return sourceSheet.getRow(rowIndex).getCell(colIndex);
    }

    private String findDescriptionFromSheet(Sheet sourceSheet, int rowIndex) {
        int colIndex = this.findColumnIndex(sourceSheet, "Description");
        return sourceSheet.getRow(rowIndex).getCell(colIndex).toString();
    }

    private Cell findTotalHoursFromSheet(Sheet sourceSheet, int rowIndex) {
        int colIndex = this.findColumnIndex(sourceSheet, "Total Hours");
        return sourceSheet.getRow(rowIndex).getCell(colIndex);
    }

    private void updateDescriptionAndHours(Sheet sourceSheet, Sheet destinationSheet, String name, CellStyle style) {
        Set<String> allTasks = new HashSet<>();
        int rowCount = sourceSheet.getPhysicalNumberOfRows();

        for (int rowIndex = 1; rowIndex < rowCount; ++rowIndex) {
            Row row = sourceSheet.getRow(rowIndex);
            if (row != null) {
                String namePresent = this.findNameFromSheet(sourceSheet, rowIndex);
                if (namePresent.equals(name)) {
                    Cell dateCell = this.findDateFromSheet(sourceSheet, rowIndex);
                    int getDay = this.getDayFromDate(dateCell.toString());
                    Cell descriptionCell = destinationSheet.getRow(getDay).getCell(3);
                    Cell projectTimeCell = destinationSheet.getRow(getDay).getCell(4);
                    String existingTask = descriptionCell.toString();
                    String newTask = this.findDescriptionFromSheet(sourceSheet, rowIndex);
                    StringBuilder token = new StringBuilder(newTask);
                    token.append("#");
                    token.append(getDay);

                    if (!allTasks.contains(token.toString())) {
                        allTasks.add(token.toString());
                        if (!existingTask.isEmpty()) {
                            existingTask = existingTask + ", ";
                        }

                        existingTask = existingTask + newTask;
                    }

                    Cell hoursCell = this.findTotalHoursFromSheet(sourceSheet, rowIndex);
                    String hours = hoursCell.getStringCellValue();
                    double hoursDouble = this.findHoursInDouble(hours);
                    projectTimeCell.setCellValue(hoursDouble);
                    projectTimeCell.setCellStyle(style);
                    descriptionCell.setCellValue(existingTask);
                    descriptionCell.setCellStyle(style);
                }
            }
        }

    }

    private int getDayFromDate(String date) {
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern(OUT_PUT_DATE_FORMAT);

        try {
            LocalDate parsedDate = LocalDate.parse(date, formatter);
            return parsedDate.getDayOfMonth();
        } catch (DateTimeParseException var4) {
            System.out.println("Invalid date format: " + date);
            throw var4;
        }
    }

    private boolean isWeekend(LocalDate date) {
        DayOfWeek dayOfWeek = date.getDayOfWeek();
        return dayOfWeek == DayOfWeek.SATURDAY || dayOfWeek == DayOfWeek.SUNDAY;
    }

    private void fillSummarySheet(Sheet destinationSheet, Map<String, Double> employeeNames) {
        int rowIndex = 1;
        int totalHours = 0;
        Iterator<String> var7 = employeeNames.keySet().iterator();

        Cell thirdCell;
        Cell totalHoursSecondCol;
        Cell blankThirdCol;
        while (var7.hasNext()) {
            String name = var7.next();
            Row row = destinationSheet.createRow(rowIndex++);
            thirdCell = row.createCell(0);
            thirdCell.setCellValue(name);
            double hours = employeeNames.get(name);
            totalHours = (int) ((double) totalHours + hours);
            totalHoursSecondCol = row.createCell(1);
            totalHoursSecondCol.setCellValue(hours);
            blankThirdCol = row.createCell(2);
            blankThirdCol.setCellValue("Existing");
        }

        Row blankRow = destinationSheet.createRow(rowIndex++);
        blankRow.createCell(0);
        blankRow.createCell(1);
        blankRow.createCell(2);
        Row totalHoursRow = destinationSheet.createRow(rowIndex);
        Cell totalHoursFirstCol = totalHoursRow.createCell(0);
        totalHoursFirstCol.setCellValue("Total Hours");
        totalHoursSecondCol = totalHoursRow.createCell(1);
        totalHoursSecondCol.setCellValue(totalHours);
        totalHoursRow.createCell(2);
    }

    private void addColumns(String[] columns, Sheet sheet) {
        Row row = sheet.createRow(0);
        int columnIndex = 0;
        for (String column : columns) {
            Cell cell = row.createCell(columnIndex++);
            cell.setCellValue(column);
        }

    }

    private int findEmployeeNameIndex(Sheet sheet) {
        Row firstRow = sheet.getRow(0);
        int columnIndex = 0;

        for (Iterator<Cell> var4 = firstRow.iterator(); var4.hasNext(); ++columnIndex) {
            Cell column = var4.next();
            if (column.toString().equals("Emp Name")) {
                return columnIndex;
            }
        }

        return -1;
    }

    private int findTotalHoursIndex(Sheet sheet) {
        Row firstRow = sheet.getRow(0);
        int columnIndex = 0;

        for (Iterator<Cell> var4 = firstRow.iterator(); var4.hasNext(); ++columnIndex) {
            Cell column = var4.next();
            if (column.toString().equals("Total Hours")) {
                return columnIndex;
            }
        }

        return -1;
    }

    private TreeMap<String, Double> findAllEmployeeNames(Sheet sheet) {
        TreeMap<String, Double> store = new TreeMap<>();
        Map<String, Set<Integer>> visited = new HashMap<>();
        int empNameIndex = this.findEmployeeNameIndex(sheet);
        int totalHoursIndex = this.findTotalHoursIndex(sheet);
        Iterator<Row> var6 = sheet.iterator();

        while (true) {
            Cell hoursCell;
            String name;
            int day;
            do {
                Row row;
                do {
                    Cell cell;
                    do {
                        do {
                            if (!var6.hasNext()) {
                                return store;
                            }

                            row = var6.next();
                            cell = row.getCell(empNameIndex);
                        } while (cell == null);

                        hoursCell = row.getCell(totalHoursIndex);
                        name = cell.toString();
                    } while (name.equals("Emp Name"));
                } while (hoursCell == null);

                Cell dateCell = row.getCell(3);
                day = this.getDayFromDate(dateCell.toString());
            } while (visited.containsKey(name) && visited.get(name).contains(day));

            String hours = hoursCell.getStringCellValue();
            double hoursDouble = this.findHoursInDouble(hours);
            store.put(name, store.getOrDefault(name, 0.0) + hoursDouble);
            if (!visited.containsKey(name)) {
                visited.put(name, new HashSet<>());
            }

            visited.get(name).add(day);
        }
    }

    private double findHoursInDouble(String hours) {
        if (hours == null) {
            return 0;
        }
        String[] timeData = hours.split(":");
        if ((timeData.length != 2)) {
            return 0;
        }
        double parsedHours = Double.parseDouble(timeData[0]);
        double minutes = Double.parseDouble(timeData[1]);
        if (minutes == 15) {
            parsedHours += 0.3;
        } else if (minutes == 30) {
            parsedHours += 0.5;
        } else if (minutes == 45) {
            parsedHours += 0.75;
        }
        return parsedHours;
    }
}
