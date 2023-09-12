package com.example.exceldemo.entity;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Data;

@Data
public class TableOne {

    @ExcelProperty("标段id")
    private String id;

    @ExcelProperty("构件编码")
    private String memberCode;

    @ExcelProperty("构件名称")
    private String memberName;

    @ExcelProperty("全路径")
    private String meberPath;
}
