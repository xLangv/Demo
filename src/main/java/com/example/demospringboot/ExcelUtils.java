package com.example.demospringboot;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.read.metadata.ReadSheet;
import com.alibaba.excel.write.handler.WriteHandler;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.style.column.AbstractColumnWidthStyleStrategy;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.springframework.stereotype.Component;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

@Component
public class ExcelUtils {
    /**
     * 通用Excel导入方法
     *
     * @param file        Excel文件
     * @param listenerOne 自定义监听器
     * @param listenerTwo 自定义监听器
     */
    public void importExcel(MultipartFile file,
                            AnalysisEventListener<?> listenerOne,
                            AnalysisEventListener<?> listenerTwo) {
        try (ExcelReader excelReader = EasyExcel.read(file.getInputStream()).build()) {
            // 创建两个 ReadSheet 对象
            ReadSheet readSheet1 = EasyExcel.readSheet("用户信息")
                    .head(UserImportDTO.class) // 第一个 Sheet 的头部类型
                    .registerReadListener(listenerOne)
                    .build();

            ReadSheet readSheet2 = EasyExcel.readSheet("用户信息机构关联")
                    .head(UserOrganDTO.class) // 第二个 Sheet 的头部类型
                    .registerReadListener(listenerTwo)
                    .build();

            // 一次性读取多个 Sheet，避免性能浪费
            excelReader.read(readSheet1, readSheet2);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    /**
     * 创建两个Sheet的导入模板
     */
    public void createImportTemplate(HttpServletResponse response) throws IOException {
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setCharacterEncoding("utf-8");
        String fileName = URLEncoder.encode("用户导入模板", "UTF-8").replaceAll("\\+", "%20");
        response.setHeader("Content-disposition", "attachment;filename*=utf-8''" + fileName + ".xlsx");

        try (ExcelWriter excelWriter = EasyExcel.write(response.getOutputStream()).build()) {

            // Sheet1 - 用户导入模板
            WriteSheet userImportSheet = EasyExcel.writerSheet(0, "用户信息").head(UserImportDTO.class)
                    // 注册下拉和样式处理器
                    .registerWriteHandler(createUserImportDropdownHandler()).build();

            // 生成用户导入模板数据
            List<UserImportDTO> userImportTemplateData = createUserImportTemplateData();
            excelWriter.write(userImportTemplateData, userImportSheet);

            // Sheet2 - 用户机构关联模板
            WriteSheet userOrganSheet = EasyExcel.writerSheet(1, "用户信息机构关联").head(UserOrganDTO.class)
                    .registerWriteHandler(createUserOrganDropdownHandler(excelWriter.writeContext().writeWorkbookHolder().getWorkbook())).build();

            // 生成用户机构关联模板数据
            List<UserOrganDTO> userOrganTemplateData = createUserOrganTemplateData();
            excelWriter.write(userOrganTemplateData, userOrganSheet);
            // 添加说明信息
            addTemplateNotes(excelWriter, userImportSheet);
        }
    }


    /**
     * 用户导入Sheet的下拉处理器
     */
    private WriteHandler createUserImportDropdownHandler() {
        return new AbstractColumnWidthStyleStrategy() {
            @Override
            protected void setColumnWidth(WriteSheetHolder writeSheetHolder, List<WriteCellData<?>> cellDataList, Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {
                Sheet sheet = writeSheetHolder.getSheet();
                String sheetName = sheet.getSheetName();
                // 性别下拉列表
                if (cell.getColumnIndex() == 4) {
                    addDropDownList1(sheet, Arrays.asList("男", "女"), relativeRowIndex + 1, 100, 4, 4, sheetName);
                } else if (cell.getColumnIndex() == 5) {// 机构类型下拉列表
                    addDropDownList1(sheet, Arrays.asList("总部", "分支机构"), relativeRowIndex + 1, 100, 5, 5, sheetName);
                }

                // 设置列宽
                sheet.setColumnWidth(cell.getColumnIndex(), 6000);
            }
        };
    }

    /**
     * 用户机构关联Sheet的下拉处理器
     */
    private WriteHandler createUserOrganDropdownHandler(Workbook workbook) {
        return new AbstractColumnWidthStyleStrategy() {
            @Override
            protected void setColumnWidth(WriteSheetHolder writeSheetHolder, List<WriteCellData<?>> cellDataList, Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {
                Sheet sheet = writeSheetHolder.getSheet();
                String sheetName = sheet.getSheetName();
                // 用户机构关联  第一列 登录名 使用用户信息的录入的登录名 实现登录名关联多个机构
                if (cell.getColumnIndex() == 0) {
                    String hiddenSheetName = "用户信息";

                    // 创建名称管理器
                    Name name = workbook.getName(hiddenSheetName);
                    if (name == null) {
                        name = workbook.createName();
                        name.setNameName(hiddenSheetName);
                    }
                    name.setRefersToFormula(hiddenSheetName + "!$A$2:$A$100");


                    // 设置数据验证
                    DataValidationHelper helper = sheet.getDataValidationHelper();
                    DataValidationConstraint constraint = helper.createFormulaListConstraint(hiddenSheetName);
                    CellRangeAddressList addressList = new CellRangeAddressList(1, 100, 0, 0);
                    DataValidation validation = helper.createValidation(constraint, addressList);

                    // 设置错误提示信息
                    validation.setShowErrorBox(true);
                    validation.setErrorStyle(DataValidation.ErrorStyle.STOP);
                    validation.createErrorBox("提示", "请从下拉列表中选择");
                    validation.setShowPromptBox(true);
                    validation.createPromptBox("提示", "请选择");

                    sheet.addValidationData(validation);
                } else if (cell.getColumnIndex() == 1) {// 机构列表下拉列表
                    addDropDownList(sheet, getOrganListOptions(), relativeRowIndex + 1, 100, 1, 1, sheetName);
                }
                // 设置列宽
                sheet.setColumnWidth(cell.getColumnIndex(), 6000);
            }
        };
    }

    /**
     * 获取机构列表选项
     */
    private List<String> getOrganListOptions() {
        // 这里替换成实际的机构数据获取逻辑
        List<String> data = new ArrayList<>();
        for (int i = 0; i < 1000; i++) {
            data.add(String.valueOf(i));
        }
        return data;
    }

    /**
     * 生成用户导入模板数据
     */
    private List<UserImportDTO> createUserImportTemplateData() {
        return new ArrayList<>();
    }

    /**
     * 生成用户机构关联模板数据
     */
    private List<UserOrganDTO> createUserOrganTemplateData() {
        return new ArrayList<>();
    }

    /**
     * 添加模板说明
     */
    private void addTemplateNotes(ExcelWriter excelWriter, WriteSheet sheet) {
        Sheet excelSheet = excelWriter.writeContext().writeSheetHolder().getSheet();
        Row noteRow = excelSheet.createRow(excelSheet.getLastRowNum() + 2);
        Cell noteCell = noteRow.createCell(0);
        noteCell.setCellValue("填写说明：\n" + "1. 用户导入Sheet：填写用户基本信息\n" + "2. 用户机构关联Sheet：选择用户，关联多个机构\n" + "3. 机构类型：1-总部，2-分支机构\n" + "4. 下拉选择后，可直接选择");

        CellStyle noteStyle = excelSheet.getWorkbook().createCellStyle();
        noteStyle.setWrapText(true);
        noteCell.setCellStyle(noteStyle);

        // 合并单元格显示说明
        excelSheet.addMergedRegion(new CellRangeAddress(noteRow.getRowNum(), noteRow.getRowNum() + 4, 0, 6));
    }

    /**
     * 添加下拉列表验证
     */
    private void addDropDownList1(Sheet sheet, List<String> list, int firstRow, int lastRow, int firstCol, int lastCol, String sheetName) {
        // 直接使用DataValidationHelper创建约束
        DataValidationHelper helper = sheet.getDataValidationHelper();
        // 直接使用字符串数组作为下拉选项，而不是引用隐藏sheet
        DataValidationConstraint constraint = helper.createExplicitListConstraint(list.toArray(new String[0]));
        CellRangeAddressList addressList = new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol);
        DataValidation validation = helper.createValidation(constraint, addressList);

        // 设置错误提示信息
        validation.setShowErrorBox(true);
        validation.setErrorStyle(DataValidation.ErrorStyle.STOP);
        validation.createErrorBox("无效输入", "请从下拉列表中选择有效的选项");
        validation.setShowPromptBox(true);
        validation.createPromptBox("选择提示", "请从下拉列表中选择");

        sheet.addValidationData(validation);
    }

    /**
     * 添加下拉列表验证
     */
    private void addDropDownList(Sheet sheet, List<String> list, int firstRow, int lastRow, int firstCol, int lastCol, String sheetName) {
        String hiddenSheetName = "hidden" + firstCol + sheetName;
        Workbook workbook = sheet.getWorkbook();

        // 创建或获取隐藏的sheet页
        Sheet hiddenSheet = workbook.getSheet(hiddenSheetName);
        if (hiddenSheet == null) {
            hiddenSheet = workbook.createSheet(hiddenSheetName);
            workbook.setSheetHidden(workbook.getSheetIndex(hiddenSheet), true);
        }

        // 在隐藏的sheet页添加下拉数据
        for (int i = 0; i < list.size(); i++) {
            Row row = hiddenSheet.getRow(i);
            if (row == null) {
                row = hiddenSheet.createRow(i);
            }
            Cell cell = row.getCell(0);
            if (cell == null) {
                cell = row.createCell(0);
            }
            cell.setCellValue(list.get(i));
        }

        // 创建名称管理器
        Name name = workbook.getName(hiddenSheetName);
        if (name == null) {
            name = workbook.createName();
            name.setNameName(hiddenSheetName);
        }
        // 设置固定100行的引用范围
        name.setRefersToFormula(hiddenSheetName + "!$A$1:$A$100");

        // 设置数据验证
        DataValidationHelper helper = sheet.getDataValidationHelper();
        DataValidationConstraint constraint = helper.createFormulaListConstraint(hiddenSheetName);
        CellRangeAddressList addressList = new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol);
        DataValidation validation = helper.createValidation(constraint, addressList);

        // 设置错误提示信息
        validation.setShowErrorBox(true);
        validation.setErrorStyle(DataValidation.ErrorStyle.STOP);
        validation.createErrorBox("无效输入", "请从下拉列表中选择有效的选项");
        validation.setShowPromptBox(true);
        validation.createPromptBox("选择提示", "请从下拉列表中选择");

        sheet.addValidationData(validation);
    }
}
