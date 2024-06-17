//package org.example;
//
//import org.apache.poi.ss.usermodel.*;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//import org.json.simple.JSONArray;
//import org.json.simple.JSONObject;
//
//import java.io.FileInputStream;
//import java.io.FileWriter;
//import java.io.IOException;
//
//public class ExcelToJsonConverter {
//  public static void main(String[] args) {
//    String excelFilePath = "/Users/yimingzhu/Downloads/ACGA.xlsx"; // Replace with your actual file path
//    String jsonFilePath = "/Users/yimingzhu/Downloads/output.json"; // Replace with your desired output file path
//
//    try (FileInputStream fis = new FileInputStream(excelFilePath);
//         Workbook workbook = new XSSFWorkbook(fis)) {
//      Sheet sheet = workbook.getSheetAt(0);
//      JSONArray jsonArray = new JSONArray();
//
//      for (Row row : sheet) {
//        JSONObject jsonObject = new JSONObject();
//        for (Cell cell : row) {
//          jsonObject.put("Column" + cell.getColumnIndex(), getCellValue(cell));
//        }
//        jsonArray.add(jsonObject);
//      }
//
//      try (FileWriter file = new FileWriter(jsonFilePath)) {
//        file.write(jsonArray.toJSONString());
//        file.flush();
//      }
//
//      System.out.println("Excel file converted to JSON successfully!");
//
//    } catch (IOException e) {
//      e.printStackTrace();
//    }
//  }
//
//  private static String getCellValue(Cell cell) {
//    switch (cell.getCellType()) {
//      case STRING:
//        return cell.getStringCellValue();
//      case NUMERIC:
//        if (DateUtil.isCellDateFormatted(cell)) {
//          return cell.getDateCellValue().toString();
//        } else {
//          return String.valueOf(cell.getNumericCellValue());
//        }
//      case BOOLEAN:
//        return String.valueOf(cell.getBooleanCellValue());
//      case FORMULA:
//        return cell.getCellFormula();
//      default:
//        return "";
//    }
//  }
//}
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
    String jsonFilePath = "/Users/yimingzhu/Downloads/ACGA.json";

    try (FileInputStream excelFile = new FileInputStream(new File(excelFilePath));
         Workbook workbook = new XSSFWorkbook(excelFile);
         FileWriter file = new FileWriter(jsonFilePath)) {

      // Assume the first sheet
      Sheet sheet = workbook.getSheetAt(0);
      JSONArray jsonArray = new JSONArray();

      // Get the headers
      Row headerRow = sheet.getRow(1); // Adjust index if necessary
      int columnCount = headerRow.getLastCellNum();
      String[] headers = new String[columnCount];
      for (int i = 0; i < columnCount; i++) {
        headers[i] = headerRow.getCell(i).toString();
      }

      // Iterate through rows
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


