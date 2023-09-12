package com.example.exceldemo.service;

import jakarta.servlet.http.HttpServletResponse;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;

public interface ExcelService {
    /**
     * 处理数据
     * @param multipartFile 文件
     */
    void handelData(MultipartFile multipartFile, HttpServletResponse response, String filterStr) throws IOException;

    void handRepeatB(MultipartFile multipartFile, HttpServletResponse response) throws IOException;

    void handMergeB(MultipartFile multipartFile, HttpServletResponse response) throws IOException;
}
