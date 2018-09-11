package com.suk.poihelper.excelhelper.util;


import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;

/**
 * Author  Administer
 * Time    2017/11/29 16:06
 * Des
 */

public class FileConvert {
    public static String convertCSVorTxt2Xls(String readinFile) {
        try {
            @SuppressWarnings("resource")
            BufferedReader readTxt = new BufferedReader(new InputStreamReader(new FileInputStream(readinFile),
                    isFileCodingInUTF8(readinFile) ? "UTF-8" : "GBK"));
            String inStr = "";

            //目标文件
            HSSFWorkbook writeWorkbook = new HSSFWorkbook();
            HSSFSheet targetSheet = writeWorkbook.createSheet();
            HSSFRow targetRow;
            HSSFCell targetCell;
            int rowIndex = 0;

            while ((inStr = readTxt.readLine()) != null) {  //读取

                targetRow = targetSheet.createRow(rowIndex);  //创建新行
                String str[] = inStr.split(",");  //得到列值
                for (int i = 0; i < str.length; i++) {
                    targetCell = targetRow.createCell(i);  //创建新列
                    targetCell.setCellValue(str[i]);  //赋值
                }
                rowIndex++;
            }

            String newFilePath = readinFile.replace(".csv", ".xls");
            FileOutputStream outputExcel = new FileOutputStream(newFilePath);
            writeWorkbook.write(outputExcel);
            outputExcel.flush();  //清空缓冲区数据
            outputExcel.close();


            System.out.println("--------------  export txt/csv ←→ xls  成功...");
            return newFilePath;
        } catch (Exception e) {
            // TODO: handle exception
            System.out.println("--------------  export txt/csv ←→ xls  运行异常：" + e);
            e.printStackTrace();
        }
        return null;
    }

    private static boolean isFileCodingInUTF8(String filePath) throws IOException {
        File file = new File(filePath);
        InputStream in = new FileInputStream(file);
        byte[] b = new byte[3];
        in.read(b);
        in.close();
        if (b[0] == -17 && b[1] == -69 && b[2] == -65)
            return true;
        else
            return false;
    }
}