package com.wixdom;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.AutoFilter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFAutoFilter;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.net.URL;

/**
 * @author Lin Qinghua
 */
public class JsonMergeXlsx {

    public void merge(String jsonPath, String jsonPtrExpr, String xlsPath, String sheetName) throws IOException {
        File jsonFile = new File(jsonPath);
        File xlsFile = new File(xlsPath);
        this.merge(jsonFile.toURI().toURL(), jsonPtrExpr, xlsFile.toURI().toURL(), sheetName);
    }

    public void merge(URL jsonUrl, String jsonPtrExpr, URL xlsUrl, String sheetName) throws IOException {
        JsonNode jsonNode = new ObjectMapper().readTree(jsonUrl).at(jsonPtrExpr);
        Sheet sheet = new XSSFWorkbook(xlsUrl.openStream()).getSheet(sheetName);
        this.merge(jsonNode, sheet);
    }

    private void merge(JsonNode jsonNode, Workbook workbook, String sheetName) {
    }

    private void merge(JsonNode jsonNode, Sheet sheet) {
        Row firstRow = sheet.getRow(sheet.getFirstRowNum());
        CellRangeAddress cellRangeAddress = new CellRangeAddress(sheet.getFirstRowNum(), sheet.getLastRowNum(), firstRow.getFirstCellNum(), firstRow.getLastCellNum());
        AutoFilter autoFilter = sheet.setAutoFilter(cellRangeAddress);
        XSSFAutoFilter xssfAutoFilter = (XSSFAutoFilter) autoFilter;
        System.out.println(sheet);
    }



}
