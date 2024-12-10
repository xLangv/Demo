package com.example.demospringboot;


import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Data;

// 用户导入DTO
@Data
public class UserImportDTO {
    @ExcelProperty("登录名")
    private String userName;

    @ExcelProperty("姓名")
    private String nickName;

    @ExcelProperty("手机号")
    private String phone;

    @ExcelProperty("电子邮箱")
    private String email;

    @ExcelProperty("性别")
    private String sex;

    @ExcelProperty("机构归属")
    private String organAuth;
}

