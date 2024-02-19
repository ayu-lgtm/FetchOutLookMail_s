package com.jbdl.fetchoutlookmail;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelFileUpdater {

    public static void main(String[] args) {
        String filePath = "C:/Users/Ayush Rastogi/Desktop/emails.xlsx";
        List<String[]> newData = getNewData(); // Your method to get new data
        try {
            updateExcelFile(filePath, newData);
            System.out.println("Data successfully added to the Excel file.");
        } catch (IOException e) {
            System.err.println("Error updating Excel file: " + e.getMessage());
            e.printStackTrace();
        }
    }

    public static void updateExcelFile(String filePath, List<String[]> newData) throws IOException {
        FileInputStream inputStream = null;
        Workbook workbook = null;
        FileOutputStream outputStream = null;
        try {
            File file = new File(filePath);
            if (!file.exists() || file.length() == 0) {
                // If the file does not exist or is empty, create a new workbook
                workbook = new XSSFWorkbook();
                Sheet sheet = workbook.createSheet("Emails");
                createHeaderRow(sheet);
            } else {
                // If the file exists and is not empty, open it and get the first sheet
                inputStream = new FileInputStream(file);
                workbook = WorkbookFactory.create(inputStream);
            }
            // Get the first sheet
            Sheet sheet = workbook.getSheetAt(0);
            // Append new data
            int rowNum = sheet.getLastRowNum() + 1;
            for (String[] data : newData) {
                Row row = sheet.createRow(rowNum++);
                for (int i = 0; i < data.length; i++) {
                    row.createCell(i).setCellValue(data[i]);
                }
            }
            // Write changes to the output file
            outputStream = new FileOutputStream(filePath);
            workbook.write(outputStream);
        } finally {
            // Close resources
            if (workbook != null) {
                workbook.close();
            }
            if (inputStream != null) {
                inputStream.close();
            }
            if (outputStream != null) {
                outputStream.close();
            }
        }
    }

    public static void createHeaderRow(Sheet sheet) {
        Row headerRow = sheet.createRow(0);
        String[] headers = {"Sender", "Receiver", "Subject", "Details", "Date", "Time"};
        for (int i = 0; i < headers.length; i++) {
            headerRow.createCell(i).setCellValue(headers[i]);
        }
    }

    public static List<String[]> getNewData() {
        // Your method to get new data
        return null;
    }
}
