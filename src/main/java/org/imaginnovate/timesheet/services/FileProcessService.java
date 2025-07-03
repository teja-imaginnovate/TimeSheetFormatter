package org.imaginnovate.timesheet.services;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.table.TableModel;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;

public class FileProcessService {

    private static final String OUTPUT_FILE_PATH = "TimeSheet Report.xlsx";

    private static final ConverterService converterService = new ConverterService();

    public boolean process(TableModel tableModel) {
        XSSFWorkbook workbook = this.tableModelToWorkbook(tableModel);
        try {
            cleanup();
            byte[] outPutBytes = converterService.generateFromCSV(workbook);
            generateOutPutFile(outPutBytes);
            return true;
        } catch (IOException e) {
            System.out.println(e.getMessage());
        }
        return false;
    }

    private static void cleanup() {
        try {
            boolean existingOutPutDeleted = Files.deleteIfExists(Paths.get(OUTPUT_FILE_PATH));
            if (existingOutPutDeleted) {
                System.out.println("Cleaning previous data..");
            }
        } catch (IOException e) {
            System.err.println(e.getMessage());
        }
    }

    private static void generateOutPutFile(byte[] outputStream) throws IOException {
        if (outputStream != null && outputStream.length > 0) {
            FileOutputStream fos = new FileOutputStream(OUTPUT_FILE_PATH);
            fos.write(outputStream);
            fos.close();
            System.out.println("Successfully generated time sheet report - " + OUTPUT_FILE_PATH);
        } else {
            System.err.println("Report generation failed!");
        }
    }

    private XSSFWorkbook tableModelToWorkbook(TableModel model) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("data");

        // Create header row
        Row headerRow = sheet.createRow(0);
        for (int col = 0; col < model.getColumnCount(); col++) {
            Cell cell = headerRow.createCell(col);
            cell.setCellValue(model.getColumnName(col));
        }

        // Create data rows
        for (int row = 0; row < model.getRowCount(); row++) {
            Row excelRow = sheet.createRow(row + 1);
            for (int col = 0; col < model.getColumnCount(); col++) {
                Object value = model.getValueAt(row, col);
                Cell cell = excelRow.createCell(col);
                if (value != null) {
                    cell.setCellValue(value.toString());
                } else {
                    cell.setCellValue("");
                }
            }
        }

        // Optionally, auto-size columns
        for (int col = 0; col < model.getColumnCount(); col++) {
            sheet.autoSizeColumn(col);
        }

        return workbook;
    }

}
