package org.sakura;



import org.apache.poi.ss.usermodel.*;
import org.sakura.utils.ChineseCharacterUtil;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;

public class Main {


    private int flag=0;

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
        HashMap<String,String> hashMap=new HashMap<>();
        HashSet<String> hashSet=new HashSet<>();
        int typeNum=0;
        if (workbook!=null){
            System.out.println(workbook.getNumberOfSheets());
            Sheet sheet=main.getSheet(workbook,1);
            System.out.println(sheet.getDefaultColumnWidth());
            int lastRowDataNum=sheet.getLastRowNum()-7;
            int firstRowDataNum=sheet.getFirstRowNum()+4;
            for (int i=firstRowDataNum;i<lastRowDataNum;i++){
                Row row=sheet.getRow(i);
                Cell nameCell=row.getCell(1);
                System.out.println(nameCell.getStringCellValue());
//                int lastColumnDataNum=row.getLastCellNum()-1;
                if (!hashSet.contains(nameCell.getStringCellValue())){
                    typeNum=typeNum+1;
                    hashSet.add(nameCell.getStringCellValue());
                }
            }
            Iterator<String> iterator=hashSet.iterator();
            while (iterator.hasNext()){
                String target=iterator.next();
                for(int j = firstRowDataNum;j<lastRowDataNum;j++){
                    int flag=1;
                    String STR="JKC(D)-JT-%s-%s-%s";
                    Row row=sheet.getRow(j);
                    Cell nameCell=row.getCell(1);
                    Cell dateCell=row.getCell(16);
                    if (target.equals(nameCell.getStringCellValue())){
                        String nameValue=ChineseCharacterUtil.convertHanzi2Pinyin(nameCell.getStringCellValue(),false).toUpperCase();
                        String serializeNum=String.valueOf(flag);
                        if (serializeNum.toCharArray().length<2){
                            serializeNum="000"+serializeNum;
                        }else {
                            serializeNum="00"+serializeNum;
                        }
                        System.out.println(dateCell.getCellType());
                        String tmpVal=String.valueOf(dateCell.getNumericCellValue());
                        tmpVal=tmpVal.substring(0,1)+tmpVal.substring(2,9);

                        System.out.println();
                        STR=String.format(STR,nameValue,tmpVal,serializeNum);
                    }

                    Cell resCell=row.getCell(3);
                    if (resCell!=null){
                        resCell.setCellValue(STR);
                    }else {
                        Cell cell=row.createCell(3);
                        cell.setCellType(CellType.STRING);
                        cell.setCellValue(STR);
                    }

                }
            }


            try {
                OutputStream fileOut = new FileOutputStream("/home/david/Idea Project/excel/src/main/resources/放映设备台账1.xls");
                workbook.write(fileOut);
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }


//            System.out.println(row.getFirstCellNum());
//            System.out.println();
//            System.out.println(row.getCell(20));
        }
    }
}
