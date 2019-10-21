package org.sakura;



import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.*;

public class Main {

    private static String STR="JKC(D)-JT-%s-%s-%s";

    private Workbook readExcel(String pathName){
        InputStream inputStream= null;
        Workbook workbook=null;
        try {
            inputStream = new FileInputStream(pathName);
            workbook=WorkbookFactory.create(inputStream);
            return workbook;
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }


    private Sheet getSheet(Workbook workbook,int sheetNum){
        return workbook.getSheetAt(sheetNum);
    }

    public static void main(String[] args) {
        Main main=new Main();
        Workbook workbook=main.readExcel("/home/david/Idea Project/excel/src/main/resources/放映设备台账.xls");
        if (workbook!=null){
            System.out.println(workbook.getNumberOfSheets());
            Sheet sheet=main.getSheet(workbook,1);
            System.out.println(sheet.getDefaultColumnWidth());
            int lastRowDataNum=sheet.getLastRowNum()-7;
            int firstRowDataNum=sheet.getFirstRowNum()+4;
            for (int i=firstRowDataNum;i<lastRowDataNum;i++){
                Row row=sheet.getRow(4);
                int lastColumnDataNum=row.getLastCellNum()-1;
            }




//            System.out.println(row.getFirstCellNum());
//            System.out.println();
//            System.out.println(row.getCell(20));
        }
    }
}
