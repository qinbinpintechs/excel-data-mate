package com.example.exceldemo.entity;


import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Data;

@Data
public class TableTwo {

    @ExcelProperty("标段id")
    private String id;

    @ExcelProperty("glid")
    private String glId;

    @ExcelProperty("externalId")
    private String externalId;

    @ExcelProperty("value")
    private String value;

    @ExcelProperty("propertyname")
    private String propertyName;
}
