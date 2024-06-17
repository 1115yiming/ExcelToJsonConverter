package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;

import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;

public class ExcelToJsonConverter {
  public static void main(String[] args) {
    String excelFilePath = "/Users/yimingzhu/Downloads/ACGA.xlsx"; // Replace with your actual file path
    String jsonFilePath = "/Users/yimingzhu/Downloads/output.json"; // Replace with your desired output file path

    try (FileInputStream fis = new FileInputStream(excelFilePath);
         Workbook workbook = new XSSFWorkbook(fis)) {
      Sheet sheet = workbook.getSheetAt(0);
      JSONArray jsonArray = new JSONArray();

      for (Row row : sheet) {
        JSONObject jsonObject = new JSONObject();
        for (Cell cell : row) {
          jsonObject.put("Column" + cell.getColumnIndex(), getCellValue(cell));
        }
        jsonArray.add(jsonObject);
      }

      try (FileWriter file = new FileWriter(jsonFilePath)) {
        file.write(jsonArray.toJSONString());
        file.flush();
      }

      System.out.println("Excel file converted to JSON successfully!");

    } catch (IOException e) {
      e.printStackTrace();
    }
  }

  private static String getCellValue(Cell cell) {
    switch (cell.getCellType()) {
      case STRING:
        return cell.getStringCellValue();
      case NUMERIC:
        if (DateUtil.isCellDateFormatted(cell)) {
          return cell.getDateCellValue().toString();
        } else {
          return String.valueOf(cell.getNumericCellValue());
        }
      case BOOLEAN:
        return String.valueOf(cell.getBooleanCellValue());
      case FORMULA:
        return cell.getCellFormula();
      default:
        return "";
    }
  }
}

