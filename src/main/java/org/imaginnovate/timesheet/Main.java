package org.imaginnovate.timesheet;

import com.opencsv.CSVReader;
import com.opencsv.exceptions.CsvValidationException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.imaginnovate.timesheet.services.ConverterService;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Arrays;


public class Main {

    private static final ConverterService converterService = new ConverterService();
    private static final String OUTPUT_FILE_PATH = "TimeSheet Report.xlsx";

    public static void main(String[] args) {
        if (args.length != 1) {
            System.err.println("Usage: java -jar TimeSheetFormatter.jar <input file path - csv | xlsx>");
            return;
        }
        String inputFilePath = args[0];
        File inputFile = new File(inputFilePath);
        if (!inputFile.exists()) {
            System.err.println("Input file does not exist: " + inputFilePath);
            return;
        }
        cleanup();
        String fileName = inputFile.getName();
        int dotIndex = fileName.lastIndexOf('.');
        if (dotIndex == -1 || dotIndex == fileName.length() - 1) {
            System.err.println("Invalid or malformed input file - No extension found.");
            return;
        }

        String fileExtension = fileName.substring(dotIndex + 1);
        if ("csv".equals(fileExtension)) {
            System.out.println("csv file type detected...");
            try {
                XSSFWorkbook workbook = convertCSVToExcel(inputFile);
                System.out.println("CSV converted to XLSX successfully!, Generating time sheet report....");
                byte[] outputStream = converterService.generateFromCSV(workbook);
                generateOutPutFile(outputStream);
                return;
            } catch (CsvValidationException | IOException e) {
                System.err.println(e.getMessage());
                System.err.println(Arrays.toString(e.getStackTrace()));
                return;
            }
        }
        if ("xlsx".equals(fileExtension)) {
            System.out.println("xlsx file type detected");
            try {
                FileInputStream fileInputStream = new FileInputStream(inputFile);
                byte[] buffer = new byte[(int) inputFile.length()];
                fileInputStream.read(buffer);
                fileInputStream.close();
                ByteArrayInputStream byteArrayInputStream = new ByteArrayInputStream(buffer);
                byte[] outPutStream = converterService.generateFromExcel(byteArrayInputStream);
                generateOutPutFile(outPutStream);
                return;
            } catch (IOException exception) {
                System.err.println(exception.getMessage());
                System.err.println(Arrays.toString(exception.getStackTrace()));
                return;
            }
        }
        System.err.println("Invalid file type, csv | xlsx are supported.");

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

    private static XSSFWorkbook convertCSVToExcel(File inputFile) throws IOException, CsvValidationException {
        CSVReader reader = new CSVReader(new FileReader(inputFile.getPath()));
        XSSFWorkbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("data"); // Create a new sheet

        String[] lineItems;
        int rowNum = 0;

        while ((lineItems = reader.readNext()) != null) {
            Row row = sheet.createRow(rowNum++);
            for (int i = 0; i < lineItems.length; i++) {
                Cell cell = row.createCell(i);
                cell.setCellValue(lineItems[i]);
            }
        }
        return workbook;
    }
}