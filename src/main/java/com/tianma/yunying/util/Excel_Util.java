package com.tianma.yunying.util;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;

public class Excel_Util {
    public static XSSFWorkbook workbook; // 工作簿
    public static XSSFSheet sheet; // 工作表
    public static XSSFRow row; // 行
    public static XSSFCell cell; // 列
    public static DateFormat format=new SimpleDateFormat("yyyy/MM/dd");

    public static String readExcelData(String fileName, String sheetName, int rownum, int cellnum) throws Exception{
        sheet = workbook.getSheet(sheetName);
        sheet.getRow(rownum).getCell(cellnum, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellType(CellType.STRING);
        String cellValue = sheet.getRow(rownum).getCell(cellnum).getStringCellValue();
        return  cellValue;
    }

    public static int readrowNum(String fileName, String sheetName) throws Exception{
        int count = 0;
        sheet = workbook.getSheet(sheetName);
        int rowNum=sheet.getLastRowNum();
        for(int i = 1; i < rowNum; i++){
            if(readExcelData(fileName,sheetName,i,0).equals("")){
                if(readExcelData(fileName,sheetName,i,1).equals("") &&readExcelData(fileName,sheetName,i,2).equals(""))
                    count++;
            }
        }
        return  rowNum - count; //去掉空行
    }

    public static int readcolNum(String fileName, String sheetName) throws Exception{
        sheet = workbook.getSheet(sheetName);
        int columnNum=sheet.getRow(0).getPhysicalNumberOfCells();
        return  columnNum;
    }
    public static void closeExcel() throws IOException {
        workbook.close();
    }

    public static  int readWantCol(String fileName, String sheetName, int rowNum,String colName) throws Exception {
        int index = 0;
        sheet = workbook.getSheet(sheetName);
        for(int i = 0; i < readcolNum(fileName,sheetName); i++){
            if(readExcelData(fileName,sheetName,rowNum,i).equals(colName)){
                return i;
            }
        }
        return 0;
    }

    public static String DateToFormat(String date){
        return String.valueOf(format.format(org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.parseDouble(date))));
    }

}
