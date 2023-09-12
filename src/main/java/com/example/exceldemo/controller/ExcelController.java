package com.example.exceldemo.controller;


import com.example.exceldemo.service.ExcelService;
import io.swagger.v3.oas.annotations.Operation;
import jakarta.servlet.http.HttpServletResponse;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.MediaType;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;

@RestController
@RequestMapping("/test")
public class ExcelController {

    private final ExcelService excelService;

    public ExcelController(ExcelService excelService) {
        this.excelService = excelService;
    }

    @PostMapping(value = "/handel", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    @Operation(summary = "数据处理")
    public void handleData(@RequestPart("file") MultipartFile multipartFile, HttpServletResponse response, @RequestParam(value = "filterStr", required = false) String filterStr) throws IOException {

        excelService.handelData(multipartFile, response, filterStr);
    }


    @PostMapping(value = "/handRepeatB", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    @Operation(summary = "B表合并")
    public void handRepeatB(@RequestPart("file") MultipartFile multipartFile, HttpServletResponse response) throws IOException {
        excelService.handRepeatB(multipartFile, response);
    }

    @PostMapping(value = "/handMergeB", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    @Operation(summary = "B表重复数据筛选")
    public void handMergeB(@RequestPart("file") MultipartFile multipartFile, HttpServletResponse response) throws IOException {
        excelService.handMergeB(multipartFile, response);
    }

}
