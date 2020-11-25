package com.ouyang.luckysheet.demo.utils;
 
 
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.hssf.util.Region;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
 
import java.awt.*;
import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;
 
public class ExcelUtils {
 
    //基于模板导出
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
            if (jsonObjectValue != null && jsonObjectValue.get("v") != null) {
                value = jsonObjectValue.get("v") + "";
            }
            if (sheet.getRow((int) object.get("r")) != null && sheet.getRow((int) object.get("r")).getCell((int) object.get("c")) != null) {
                sheet.getRow((int) object.get("r")).getCell((int) object.get("c")).setCellValue(value);
            } else {
                System.out.println("错误的=" + index + ">>>" + str_);
            }
 
 
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
 
 
    /***
     * 基于POI解析 从0开始导出xlsx文件，不是基于模板
     * @param title 表格名
     * @param newFileDir 保存的文件夹名
     * @param newFileName 保存的文件名
     * @param excelData luckysheet 表格数据
     */
    public static void exportLuckySheetXlsxByPOI(String title, String newFileDir, String newFileName, String excelData) {
        excelData = excelData.replace("&#xA;", "\\r\\n");//去除luckysheet中 &#xA 的换行
        JSONArray jsonArray = (JSONArray) JSONObject.parse(excelData);
        for (int sheetIndex = 0; sheetIndex < jsonArray.size(); sheetIndex++) {
            JSONObject jsonObject = (JSONObject) jsonArray.get(sheetIndex);
            JSONArray celldataObjectList = jsonObject.getJSONArray("celldata");
            JSONArray rowObjectList = jsonObject.getJSONArray("visibledatarow");
            JSONArray colObjectList = jsonObject.getJSONArray("visibledatacolumn");
            JSONArray dataObjectList = jsonObject.getJSONArray("data");
            JSONObject mergeObject = jsonObject.getJSONObject("config").getJSONObject("merge");//合并单元格
            JSONObject columnlenObject = jsonObject.getJSONObject("config").getJSONObject("columnlen");//表格列宽
            JSONObject rowlenObject = jsonObject.getJSONObject("config").getJSONObject("rowlen");//表格行高
            JSONArray borderInfoObjectList = jsonObject.getJSONObject("config").getJSONArray("borderInfo");//边框样式
            //参考：https://blog.csdn.net/jdtugfcg/article/details/84100315
            //创建操作Excel的XSSFWorkbook对象
            XSSFWorkbook excel = new XSSFWorkbook();
            XSSFCellStyle cellStyle = excel.createCellStyle();
            //创建XSSFSheet对象
            XSSFSheet sheet = excel.createSheet(jsonObject.getString("name"));
 
            //我们都知道excel是表格，即由一行一行组成的，那么这一行在java类中就是一个XSSFRow对象，我们通过XSSFSheet对象就可以创建XSSFRow对象
            //如：创建表格中的第一行（我们常用来做标题的行)  XSSFRow firstRow = sheet.createRow(0); 注意下标从0开始
            //根据luckysheet创建行列
            //创建行和列
            for (int i = 0; i < rowObjectList.size(); i++) {
                XSSFRow row = sheet.createRow(i);//创建行
                try {
                    row.setHeightInPoints(Float.parseFloat(rowlenObject.get(i) + ""));//行高px值
                } catch (Exception e) {
                    row.setHeightInPoints(20f);//默认行高
                }
 
                for (int j = 0; j < colObjectList.size(); j++) {
                    if (columnlenObject.getInteger(j + "") != null) {
                        sheet.setColumnWidth(j, columnlenObject.getInteger(j + "") * 42);//列宽px值
                    }
                    row.createCell(j);//创建列
                }
            }
 
            //设置值,样式
            setCellValue(celldataObjectList, borderInfoObjectList, sheet, excel);
 
            // 判断路径是否存在
            File dir = new File(newFileDir);
            if (!dir.exists()) {
                dir.mkdirs();
            }
            OutputStream out = null;
            try {
                out = new FileOutputStream(newFileDir + newFileName);
 
                excel.write(out);
 
                out.close();
 
            } catch (FileNotFoundException e) {
                e.printStackTrace();
 
            } catch (IOException e) {
                e.printStackTrace();
 
            }
        }
 
 
    }
 
 
    private static void setMergeAndColorByObject(com.alibaba.fastjson.JSONObject jsonObjectValue, XSSFSheet sheet, XSSFCellStyle style) {
        JSONObject mergeObject = (JSONObject) jsonObjectValue.get("mc");
        if (mergeObject != null) {
            int r = (int) (mergeObject.get("r"));
            int c = (int) (mergeObject.get("c"));
            if ((mergeObject.get("rs") != null && (mergeObject.get("cs") != null))) {
                int rs = (int) (mergeObject.get("rs"));
                int cs = (int) (mergeObject.get("cs"));
                CellRangeAddress region = new CellRangeAddress(r, r + rs - 1, (short) (c), (short) (c + cs - 1));
                sheet.addMergedRegion(region);
            }
        }
 
        if (jsonObjectValue.getString("bg") != null) {
            int bg = Integer.parseInt(jsonObjectValue.getString("bg").replace("#", ""), 16);
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);    //设置填充方案
            style.setFillForegroundColor(new XSSFColor(new Color(bg)));  //设置填充颜色
        }
 
    }
 
    private static void setBorder(JSONArray borderInfoObjectList, XSSFWorkbook workbook, XSSFSheet sheet) {
        //设置边框样式map
        Map<Integer, BorderStyle> bordMap = new HashMap<>();
        bordMap.put(1, BorderStyle.THIN);
        bordMap.put(2, BorderStyle.HAIR);
        bordMap.put(3, BorderStyle.DOTTED);
        bordMap.put(4, BorderStyle.DASHED);
        bordMap.put(5, BorderStyle.DASH_DOT);
        bordMap.put(6, BorderStyle.DASH_DOT_DOT);
        bordMap.put(7, BorderStyle.DOUBLE);
        bordMap.put(8, BorderStyle.MEDIUM);
        bordMap.put(9, BorderStyle.MEDIUM_DASHED);
        bordMap.put(10, BorderStyle.MEDIUM_DASH_DOT);
        bordMap.put(11, BorderStyle.MEDIUM_DASH_DOT_DOTC);
        bordMap.put(12, BorderStyle.SLANTED_DASH_DOT);
        bordMap.put(13, BorderStyle.THICK);
 
        //一定要通过 cell.getCellStyle()  不然的话之前设置的样式会丢失
        //设置边框
        for (int i = 0; i < borderInfoObjectList.size(); i++) {
            JSONObject borderInfoObject = (JSONObject) borderInfoObjectList.get(i);
            if (borderInfoObject.get("rangeType").equals("cell")) {//单个单元格
                JSONObject borderValueObject = borderInfoObject.getJSONObject("value");
 
                JSONObject l = borderValueObject.getJSONObject("l");
                JSONObject r = borderValueObject.getJSONObject("r");
                JSONObject t = borderValueObject.getJSONObject("t");
                JSONObject b = borderValueObject.getJSONObject("b");
 
 
                int row = borderValueObject.getInteger("row_index");
                int col = borderValueObject.getInteger("col_index");
 
                XSSFCell cell = sheet.getRow(row).getCell(col);
 
 
                if (l != null) {
                    cell.getCellStyle().setBorderLeft(bordMap.get((int) l.get("style"))); //左边框
                    int bg = Integer.parseInt(l.getString("color").replace("#", ""), 16);
                    cell.getCellStyle().setLeftBorderColor(new XSSFColor(new Color(bg)));//左边框颜色
                }
                if (r != null) {
                    cell.getCellStyle().setBorderRight(bordMap.get((int) r.get("style"))); //右边框
                    int bg = Integer.parseInt(r.getString("color").replace("#", ""), 16);
                    cell.getCellStyle().setRightBorderColor(new XSSFColor(new Color(bg)));//右边框颜色
                }
                if (t != null) {
                    cell.getCellStyle().setBorderTop(bordMap.get((int) t.get("style"))); //顶部边框
                    int bg = Integer.parseInt(t.getString("color").replace("#", ""), 16);
                    cell.getCellStyle().setTopBorderColor(new XSSFColor(new Color(bg)));//顶部边框颜色
                }
                if (b != null) {
                    cell.getCellStyle().setBorderBottom(bordMap.get((int) b.get("style"))); //底部边框
                    int bg = Integer.parseInt(b.getString("color").replace("#", ""), 16);
                    cell.getCellStyle().setBottomBorderColor(new XSSFColor(new Color(bg)));//底部边框颜色
                }
            } else if (borderInfoObject.get("rangeType").equals("range")) {//选区
                int bg_ = Integer.parseInt(borderInfoObject.getString("color").replace("#", ""), 16);
                int style_ = borderInfoObject.getInteger("style");
 
                JSONObject rangObject = (JSONObject) ((JSONArray) (borderInfoObject.get("range"))).get(0);
 
                JSONArray rowList = rangObject.getJSONArray("row");
                JSONArray columnList = rangObject.getJSONArray("column");
 
 
                for (int row_ = rowList.getInteger(0); row_ < rowList.getInteger(rowList.size() - 1) + 1; row_++) {
                    for (int col_ = columnList.getInteger(0); col_ < columnList.getInteger(columnList.size() - 1) + 1; col_++) {
                        XSSFCell cell = sheet.getRow(row_).getCell(col_);
 
                        cell.getCellStyle().setBorderLeft(bordMap.get(style_)); //左边框
                        cell.getCellStyle().setLeftBorderColor(new XSSFColor(new Color(bg_)));//左边框颜色
                        cell.getCellStyle().setBorderRight(bordMap.get(style_)); //右边框
                        cell.getCellStyle().setRightBorderColor(new XSSFColor(new Color(bg_)));//右边框颜色
                        cell.getCellStyle().setBorderTop(bordMap.get(style_)); //顶部边框
                        cell.getCellStyle().setTopBorderColor(new XSSFColor(new Color(bg_)));//顶部边框颜色
                        cell.getCellStyle().setBorderBottom(bordMap.get(style_)); //底部边框
                        cell.getCellStyle().setBottomBorderColor(new XSSFColor(new Color(bg_)));//底部边框颜色 }
                    }
                }
 
 
            }
        }
    }
 
    private static void setCellValue(JSONArray jsonObjectList, JSONArray borderInfoObjectList, XSSFSheet
            sheet, XSSFWorkbook workbook) {
        //设置字体大小和颜色
        Map<Integer, String> fontMap = new HashMap<>();
        fontMap.put(-1, "Arial");
        fontMap.put(0, "Times New Roman");
        fontMap.put(1, "Arial");
        fontMap.put(2, "Tahoma");
        fontMap.put(3, "Verdana");
        fontMap.put(4, "微软雅黑");
        fontMap.put(5, "宋体");
        fontMap.put(6, "黑体");
        fontMap.put(7, "楷体");
        fontMap.put(8, "仿宋");
        fontMap.put(9, "新宋体");
        fontMap.put(10, "华文新魏");
        fontMap.put(11, "华文行楷");
        fontMap.put(12, "华文隶书");
 
        for (int index = 0; index < jsonObjectList.size(); index++) {
            XSSFCellStyle style = workbook.createCellStyle();//样式
            XSSFFont font = workbook.createFont();//字体样式
 
            com.alibaba.fastjson.JSONObject object = jsonObjectList.getJSONObject(index);
            String str_ = (int) object.get("r") + "_" + object.get("c") + "=" + ((com.alibaba.fastjson.JSONObject) object.get("v")).get("v") + "\n";
            JSONObject jsonObjectValue = ((com.alibaba.fastjson.JSONObject) object.get("v"));
 
            String value = "";
            if (jsonObjectValue != null && jsonObjectValue.get("v") != null) {
                value = jsonObjectValue.getString("v");
            }
 
            if (sheet.getRow((int) object.get("r")) != null && sheet.getRow((int) object.get("r")).getCell((int) object.get("c")) != null) {
                XSSFCell cell = sheet.getRow((int) object.get("r")).getCell((int) object.get("c"));
                if (jsonObjectValue != null && jsonObjectValue.get("f") != null) {//如果有公式，设置公式
                    value = jsonObjectValue.getString("f");
                    cell.setCellFormula(value.substring(1,value.length()));//不需要=符号
                }
                //合并单元格与填充单元格颜色
                setMergeAndColorByObject(jsonObjectValue, sheet, style);
                //填充值
                cell.setCellValue(value);
                XSSFRow row = sheet.getRow((int) object.get("r"));
 
                //设置垂直水平对齐方式
                int vt = jsonObjectValue.getInteger("vt") == null ? 1 : jsonObjectValue.getInteger("vt");//垂直对齐	 0 中间、1 上、2下
                int ht = jsonObjectValue.getInteger("ht") == null ? 1 : jsonObjectValue.getInteger("ht");//0 居中、1 左、2右
                switch (vt) {
                    case 0:
                        style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
                        break;
                    case 1:
                        style.setVerticalAlignment(XSSFCellStyle.VERTICAL_TOP);
                        break;
                    case 2:
                        style.setVerticalAlignment(XSSFCellStyle.VERTICAL_BOTTOM);
                        break;
                }
                switch (ht) {
                    case 0:
                        style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
                        break;
                    case 1:
                        style.setAlignment(XSSFCellStyle.ALIGN_LEFT);
                        break;
                    case 2:
                        style.setAlignment(XSSFCellStyle.ALIGN_RIGHT);
                        break;
                }
 
 
                //设置合并单元格的样式有问题
                String ff = jsonObjectValue.getString("ff");//0 Times New Roman、 1 Arial、2 Tahoma 、3 Verdana、4 微软雅黑、5 宋体（Song）、6 黑体（ST Heiti）、7 楷体（ST Kaiti）、 8 仿宋（ST FangSong）、9 新宋体（ST Song）、10 华文新魏、11 华文行楷、12 华文隶书
                int fs = jsonObjectValue.getInteger("fs") == null ? 14 : jsonObjectValue.getInteger("fs");//字体大小
                int bl = jsonObjectValue.getInteger("bl") == null ? 0 : jsonObjectValue.getInteger("bl");//粗体	0 常规 、 1加粗
                int it = jsonObjectValue.getInteger("it") == null ? 0 : jsonObjectValue.getInteger("it");//斜体	0 常规 、 1 斜体
                String fc = jsonObjectValue.getString("fc") == null ? "" : jsonObjectValue.getString("fc");//字体颜色
                font.setFontName(fontMap.get(ff));//字体名字
 
 
                if (fc.length() > 0) {
                    font.setColor(new XSSFColor(new Color(Integer.parseInt(fc.replace("#", ""), 16))));
                }
                font.setFontName(ff);//字体名字
                font.setFontHeightInPoints((short) fs);//字体大小
                if (bl == 1) {
                    font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);//粗体显示
                }
                font.setItalic(it == 1 ? true : false);//斜体
 
 
                style.setFont(font);
                style.setWrapText(true);//设置自动换行
                cell.setCellStyle(style);
 
            } else {
                System.out.println("错误的=" + index + ">>>" + str_);
            }
 
 
        }
        //设置边框
        setBorder(borderInfoObjectList, workbook, sheet);
 
    }
 
 
}
 
 
