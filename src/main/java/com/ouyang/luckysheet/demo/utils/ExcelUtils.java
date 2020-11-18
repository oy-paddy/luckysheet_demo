package com.ouyang.luckysheet.demo.utils;


import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class ExcelUtils {


    public static void exportLuckySheetXlsx(String title, String newFileDir, String newFileName, String excelData) {

        JSONArray jsonArray = (JSONArray) JSONObject.parse(excelData);
        JSONObject jsonObject = (JSONObject) jsonArray.get(0);
        JSONArray jsonObjectList = jsonObject.getJSONArray("celldata");

        //excel模板路径
        String filePath = "file/"   + "模板.xlsx";
//        String filePath = "/Users/ouyang/Downloads/uploadTestProductFile/生产日报表.xlsx";
        File file = new File(filePath);
        FileInputStream in = null;
        try {
            in = new FileInputStream(file);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        //读取excel模板
        XSSFWorkbook wb = null;
        try {
            wb = new XSSFWorkbook(in);
        } catch (IOException e) {
            e.printStackTrace();
        }
        //读取了模板内所有sheet内容
        XSSFSheet sheet = wb.getSheetAt(0);
        //如果这行没有了，整个公式都不会有自动计算的效果的
        sheet.setForceFormulaRecalculation(true);

        for (int index = 0; index < jsonObjectList.size(); index++) {
            com.alibaba.fastjson.JSONObject object = jsonObjectList.getJSONObject(index);
            String str_ = (int) object.get("r") + "_" + object.get("c") + "=" + ((com.alibaba.fastjson.JSONObject) object.get("v")).get("v") + "\n";
            JSONObject jsonObjectValue = ((com.alibaba.fastjson.JSONObject) object.get("v"));

            String value = "";
            if (jsonObjectValue != null && jsonObjectValue.get("v") != null)
                value = jsonObjectValue.get("v") + "";
            if (sheet.getRow((int) object.get("r")) != null && sheet.getRow((int) object.get("r")).getCell((int) object.get("c")) != null)
                sheet.getRow((int) object.get("r")).getCell((int) object.get("c")).setCellValue(value);
            else
                System.out.println("错误的=" + index + ">>>" + str_);


        }

        // 保存文件的路径
//        String realPath = "/Users/ouyang/Downloads/uploadTestProductFile/其他文件列表/生产日报表/";
        // 判断路径是否存在
        File dir = new File(newFileDir);
        if (!dir.exists()) {
            dir.mkdirs();
        }
        //修改模板内容导出新模板
        FileOutputStream out = null;
        try {
            out = new FileOutputStream(newFileDir + newFileName);
            wb.write(out);
            out.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        System.out.println("生成文件成功："+newFileDir+newFileName);
    }



}


