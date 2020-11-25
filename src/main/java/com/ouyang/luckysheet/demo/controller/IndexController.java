package com.ouyang.luckysheet.demo.controller;

import com.ouyang.luckysheet.demo.utils.ExcelUtils;
import org.springframework.stereotype.Controller;
import org.springframework.ui.ModelMap;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;


import javax.servlet.http.HttpServletRequest;
import java.io.*;
import java.security.GeneralSecurityException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Properties;



@Controller
public class IndexController {

    @PostMapping("excel/downfile")
    @ResponseBody
    //http://localhost/excel/uploadData
    public String downExcelFile(RedirectAttributes redirectAttributes, @RequestParam("exceldatas") String exceldata, @RequestParam(value = "id", defaultValue = "0") int id, @RequestParam(value = "title", defaultValue = "其他") String title) {
        if (title.contains("."))
            title = title.substring(0, title.indexOf("."));
        String fileDir = "file";
        String fileDirNew = "file" + "/" + title + "/";

        Date date = new Date();
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMdd_HHmmss");
        String fileNameNew = title + "_" + dateFormat.format(date) + ".xlsx";

        //1-基于模板 导出生成文件
        //ExcelUtils.exportLuckySheetXlsx(title, fileDirNew, fileNameNew, exceldata);
        
        //2-基于POI 生成导出文件
        ExcelUtils.exportLuckySheetXlsxByPOI(title, fileDirNew, fileNameNew, exceldata);

        return "ok";
    }

    //在线上传xlsx
//    @LoginToken
    @PostMapping("excel/uploadData")
    @ResponseBody
    //http://localhost/excel/uploadData
    public String uploadExcelData(HttpServletRequest request, ModelMap modelMap, @RequestParam("exceldatas") String exceldata, @RequestParam("id") int id, @RequestParam("title") String title) {

        return "exceldata";
    }


    @GetMapping("/")
    public String index(){
        return "index";
    }



}
