package org.example;

//import java.io.*;
//import org.apache.poi.ss.usermodel.*;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//import org.json.simple.JSONArray;
//import org.json.simple.JSONObject;
//
//public class ExcelToJsonConverter {
//  public static void main(String[] args) {
//    // Paths to the Excel input file and JSON output file
//    String excelFilePath = "/Users/yimingzhu/Downloads/file.xlsx";
//    String jsonFilePath = "/Users/yimingzhu/Downloads/file.json";
//
//    try (FileInputStream excelFile = new FileInputStream(new File(excelFilePath));
//         Workbook workbook = new XSSFWorkbook(excelFile);
//         FileWriter file = new FileWriter(jsonFilePath)) {
//
//      // Assume the first sheet
//      Sheet sheet = workbook.getSheetAt(0);
//      JSONArray jsonArray = new JSONArray();
//
//      // Get the headers from the first row (index 0)
//      Row headerRow = sheet.getRow(0);
//      int columnCount = headerRow.getLastCellNum();
//      String[] headers = new String[columnCount];
//      for (int i = 0; i < columnCount; i++) {
//        headers[i] = headerRow.getCell(i).toString();
//      }
//
//      // Iterate through rows starting from the second row (index 1)
//      for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
//        Row row = sheet.getRow(rowIndex);
//        JSONObject rowData = new JSONObject();
//        for (int colIndex = 0; colIndex < columnCount; colIndex++) {
//          Cell cell = row.getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
//          rowData.put(headers[colIndex], cell.toString());
//        }
//        jsonArray.add(rowData);
//      }
//
//      // Convert JSON array to string and remove the unnecessary escape characters
//      String jsonString = jsonArray.toJSONString();
//      jsonString = jsonString.replace("\\/", "/");
//
//      // Write JSON to file
//      file.write(jsonString); // changed from file.write(jsonArray.toJSONString())
//      file.flush();
//      System.out.println("JSON file created: " + jsonFilePath);
//    } catch (IOException e) {
//      e.printStackTrace();
//    }
//  }
//}

//import java.io.*;
//import java.nio.file.*;
//import java.util.Iterator;
//import java.util.List;
//import java.util.Map;
//
//import org.apache.poi.ss.usermodel.*;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//import org.json.simple.JSONAware;
//import org.json.simple.JSONObject;
//import org.json.simple.JSONStreamAware;
//import org.json.simple.parser.JSONParser;
//import org.json.simple.parser.ParseException;
//
//public class ExcelToJsonConverter {
//  public static void main(String[] args) {
//    // Paths to the Excel input file and JSON output directory
//    String excelFilePath = "/Users/yimingzhu/Downloads/ACGA3.xlsx";
//    String jsonOutputDirectory = "/Users/yimingzhu/Downloads/";
//
//    try (FileInputStream excelFile = new FileInputStream(new File(excelFilePath));
//         Workbook workbook = new XSSFWorkbook(excelFile)) {
//
//      // Iterate through all sheets (tabs) in the Excel file
//      for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
//        Sheet sheet = workbook.getSheetAt(sheetIndex);
//        JSONArray jsonArray = new JSONArray();
//
//        // Get the headers from the first row (index 0)
//        Row headerRow = sheet.getRow(0);
//        int columnCount = headerRow.getLastCellNum();
//        String[] headers = new String[columnCount];
//        for (int i = 0; i < columnCount; i++) {
//          headers[i] = headerRow.getCell(i).toString();
//        }
//
//        // Iterate through rows starting from the second row (index 1)
//        for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
//          Row row = sheet.getRow(rowIndex);
//          JSONObject rowData = new JSONObject();
//          for (int colIndex = 0; colIndex < columnCount; colIndex++) {
//            Cell cell = row.getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
//            rowData.put(headers[colIndex], cell.toString());
//          }
//          jsonArray.add(rowData);
//        }
//
//        // Convert JSON array to pretty-printed string
//        String jsonString = jsonArray.toJSONString();
//        jsonString = prettyPrintJSON(jsonString);
//
//        // Write JSON to file
//        String sheetName = workbook.getSheetName(sheetIndex);
//        Path jsonFilePath = Paths.get(jsonOutputDirectory, sheetName + ".json");
//        Files.write(jsonFilePath, jsonString.getBytes());
//        System.out.println("JSON file created: " + jsonFilePath.toString());
//      }
//    } catch (IOException | ParseException e) {
//      e.printStackTrace();
//    }
//  }
//
//  private static String prettyPrintJSON(String jsonString) throws ParseException {
//    JSONParser parser = new JSONParser();
//    JSONArray jsonArray = (JSONArray) parser.parse(jsonString);
//    return jsonArray.toJSONStringWithIndent(4); // Custom method for pretty printing
//  }
//}
//
//// Custom method to pretty print JSON
//class JSONArray extends org.json.simple.JSONArray {
//  public String toJSONStringWithIndent(int indent) {
//    return toJSONString(indent);
//  }
//
//  private String toJSONString(int indent) {
//    StringWriter writer = new StringWriter();
//    try {
//      writeJSONString(writer, indent);
//    } catch (IOException e) {
//      // This should never happen with StringWriter
//      throw new RuntimeException(e);
//    }
//    return writer.toString();
//  }
//
//  private void writeJSONString(Writer out, int indent) throws IOException {
//    JSONValue.writeJSONString(this, out, indent);
//  }
//}
//
//class JSONValue extends org.json.simple.JSONValue {
//  public static void writeJSONString(Object value, Writer out, int indent) throws IOException {
//    if (value == null) {
//      out.write("null");
//      return;
//    }
//    if (value instanceof String) {
//      out.write("\"");
//      out.write(escape((String) value));
//      out.write("\"");
//      return;
//    }
//    if (value instanceof Double) {
//      if (((Double) value).isInfinite() || ((Double) value).isNaN())
//        out.write("null");
//      else
//        out.write(value.toString());
//      return;
//    }
//    if (value instanceof Float) {
//      if (((Float) value).isInfinite() || ((Float) value).isNaN())
//        out.write("null");
//      else
//        out.write(value.toString());
//      return;
//    }
//    if (value instanceof Number) {
//      out.write(value.toString());
//      return;
//    }
//    if (value instanceof Boolean) {
//      out.write(value.toString());
//      return;
//    }
//    if (value instanceof JSONStreamAware) {
//      ((JSONStreamAware) value).writeJSONString(out);
//      return;
//    }
//    if (value instanceof JSONAware) {
//      out.write(((JSONAware) value).toJSONString());
//      return;
//    }
//    if (value instanceof Map) {
//      JSONObject.writeJSONString((Map) value, out);
//      return;
//    }
//    if (value instanceof List) {
//      writeJSONString((List) value, out, indent);
//      return;
//    }
//    out.write(value.toString());
//  }
//
//  public static void writeJSONString(List list, Writer out, int indent) throws IOException {
//    if (list == null) {
//      out.write("null");
//      return;
//    }
//
//    boolean first = true;
//    int newIndent = indent + 4;
//    Iterator iter = list.iterator();
//
//    out.write("[\n");
//    while (iter.hasNext()) {
//      if (first)
//        first = false;
//      else
//        out.write(",\n");
//      for (int i = 0; i < newIndent; i++) {
//        out.write(" ");
//      }
//      Object value = iter.next();
//      if (value == null) {
//        out.write("null");
//        continue;
//      }
//      writeJSONString(value, out, newIndent);
//    }
//    out.write("\n");
//    for (int i = 0; i < indent; i++) {
//      out.write(" ");
//    }
//    out.write("]");
//  }
//}
import java.io.*;
import java.nio.file.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.ObjectWriter;

public class ExcelToJsonConverter {
  public static void main(String[] args) {
    // Paths to the Excel input file and JSON output directory
    String excelFilePath = "/Users/yimingzhu/Downloads/yourfile.xlsx";
    String jsonOutputDirectory = "/Users/yimingzhu/Downloads/";

    try (FileInputStream excelFile = new FileInputStream(new File(excelFilePath));
         Workbook workbook = new XSSFWorkbook(excelFile)) {

      // Iterate through all sheets (tabs) in the Excel file
      for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        JSONArray jsonArray = new JSONArray();

        // Get the headers from the first row (index 0)
        Row headerRow = sheet.getRow(0);
        int columnCount = headerRow.getLastCellNum();
        String[] headers = new String[columnCount];
        for (int i = 0; i < columnCount; i++) {
          headers[i] = headerRow.getCell(i).toString();
        }

        // Iterate through rows starting from the second row (index 1)
        for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
          Row row = sheet.getRow(rowIndex);
          JSONObject rowData = new JSONObject();
          for (int colIndex = 0; colIndex < columnCount; colIndex++) {
            Cell cell = row.getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
            rowData.put(headers[colIndex], cell.toString());
          }
          jsonArray.add(rowData);
        }

        // Convert JSON array to pretty-printed string using Jackson
        ObjectMapper mapper = new ObjectMapper();
        ObjectWriter writer = mapper.writerWithDefaultPrettyPrinter();
        String jsonString = writer.writeValueAsString(jsonArray);

        // Write JSON to file
        String sheetName = workbook.getSheetName(sheetIndex);
        Path jsonFilePath = Paths.get(jsonOutputDirectory, sheetName + ".json");
        Files.write(jsonFilePath, jsonString.getBytes());
        System.out.println("JSON file created: " + jsonFilePath.toString());
      }
    } catch (IOException e) {
      e.printStackTrace();
    }
  }
}
