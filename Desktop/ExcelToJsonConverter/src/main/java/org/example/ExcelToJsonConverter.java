package org.example;

import java.io.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;

public class ExcelToJsonConverter {
  public static void main(String[] args) {
    // Paths to the Excel input file and JSON output file
    String excelFilePath = "/Users/yimingzhu/Downloads/ACGA.xlsx";
    String jsonFilePath = "/Users/yimingzhu/Downloads/ACGA2.json";

    try (FileInputStream excelFile = new FileInputStream(new File(excelFilePath));
         Workbook workbook = new XSSFWorkbook(excelFile);
         FileWriter file = new FileWriter(jsonFilePath)) {

      // Assume the first sheet
      Sheet sheet = workbook.getSheetAt(0);
      JSONArray jsonArray = new JSONArray();

      // Get the headers from the second row (index 1)
      Row headerRow = sheet.getRow(1); // Adjust index if necessary
      int columnCount = headerRow.getLastCellNum();
      String[] headers = new String[columnCount];
      for (int i = 0; i < columnCount; i++) {
        headers[i] = headerRow.getCell(i).toString();
      }

      // Iterate through rows starting from the third row (index 2)
      for (int rowIndex = 2; rowIndex <= sheet.getLastRowNum(); rowIndex++) { // Adjust index if necessary
        Row row = sheet.getRow(rowIndex);
        JSONObject rowData = new JSONObject();
        for (int colIndex = 0; colIndex < columnCount; colIndex++) {
          Cell cell = row.getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
          rowData.put(headers[colIndex], cell.toString());
        }
        jsonArray.add(rowData);
      }

      // Write JSON to file
      file.write(jsonArray.toJSONString());
      System.out.println("JSON file created: " + jsonFilePath);
    } catch (IOException e) {
      e.printStackTrace();
    }
  }
}

