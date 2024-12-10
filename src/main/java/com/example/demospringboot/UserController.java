package com.example.demospringboot;

import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;

@RestController
@Slf4j
public class UserController {
    @Autowired
    ExcelUtils excelUtils;

    /**
     * 下载模板
     */
    @GetMapping("/template/download")
    public void downloadTemplate(HttpServletResponse response) {
        try {
            excelUtils.createImportTemplate(response);
        } catch (Exception e) {
            log.error("下载模板失败", e);
        }
    }
}
