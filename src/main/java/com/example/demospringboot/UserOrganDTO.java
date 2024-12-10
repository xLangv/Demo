package com.example.demospringboot;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Data;

@Data
public class UserOrganDTO {
    @ExcelProperty("登录名")
    private String userName;

    @ExcelProperty("机构名称")
    private String organName;
}