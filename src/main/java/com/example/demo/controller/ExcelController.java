package com.example.demo.controller;

import com.example.demo.service.ExcelService;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

@RestController
@RequestMapping("/file")
public class ExcelController {
    @Autowired
    ExcelService excelService;

    @GetMapping("/excel")
    public HttpServletResponse test(HttpServletRequest request, HttpServletResponse response) throws Exception {

        Workbook workbook = excelService.create();
        excelService.downloadx(workbook, response);
        return response;
    }

}
