package com.excel.tool.core;

import com.alibaba.fastjson.JSONObject;
import com.excel.tool.core.config.ExcelExportConfig;
import com.excel.tool.core.style.ExcelStyle;
import com.google.gson.Gson;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;

@Slf4j
public class ExportTools<T> {

    private String excelPath = "export";
    private int curRow = 0;

    // 导出excel
    public String generateExcel(ExcelExportConfig<T> excelExportConfig) throws Exception {
        Workbook workbook = new XSSFWorkbook();
        // 判断传入的配置类是否符合要求
        if (excelExportConfig.getTitles() == null || excelExportConfig.getTitles().length == 0){
            throw new Exception("请传入正确的表头");
        }
        int totalColLen = excelExportConfig.getTitles().length;
        Sheet activeSheet = workbook.createSheet(excelExportConfig.getSheetName()); // 创建工作簿
        // 判断是否需要表头
        if (excelExportConfig.isShowHeader()){
            // 创建表头
            createHeader(workbook,activeSheet,excelExportConfig.getHeader(),totalColLen);
        }
        // 创建标题
        createTitles(workbook,activeSheet,excelExportConfig.getTitles());
        // 填充数据
        fillData(activeSheet,excelExportConfig.getDataSource(),excelExportConfig.getFields());
        // 自动缩放
        autoSizeColumWidth(activeSheet,totalColLen);
        // 保存文件
        String saveName = saveFile(workbook,excelExportConfig.getExcelFileName());
        log.info("保存后的文件名为:" + saveName);
        return saveName;
    }

    /**
     * 创建excel的表头
     * @param workbook
     * @param sheet
     * @param header
     * @param mergerdCellCount
     * @return
     */
    private void createHeader(Workbook workbook, Sheet sheet, String header, int mergerdCellCount){
        Row headerRow = sheet.createRow(curRow++); // 创建行
        // 合并列
        if (mergerdCellCount > 1)
            sheet.addMergedRegion(new CellRangeAddress(0,0,0,mergerdCellCount - 1));
        Cell headerCell = headerRow.createCell(0); // 创建单元格
        headerCell.setCellStyle(ExcelStyle.getHeaderCellStyle(sheet));
        headerCell.setCellValue(header); // 赋值
        CellUtil.setAlignment(headerCell,HorizontalAlignment.CENTER);
    }

    /**
     * 创建表头
     * @param workbook
     * @param sheet
     * @param titles
     */
    private void createTitles(Workbook workbook,Sheet sheet,String[] titles){
        Row titleRow = sheet.createRow(curRow++);
        int len = titles.length;
        for (int i = 0; i < len; i++){
            Cell titleCell = titleRow.createCell(i);
            titleCell.setCellStyle(ExcelStyle.getTitleCellStyle(sheet));
            titleCell.setCellValue(titles[i]);
        }
    }

    /**
     * 填充业务数据
     * @param sheet
     * @param dataList
     */
    private void fillData(Sheet sheet, List<T> dataList,String[] fields){
        if (dataList != null){
            Gson gson = new Gson();
            int len = dataList.size();
            for (int i = 0; i < len; i++){
                T obj = dataList.get(i);
                String json = gson.toJson(obj);
                JSONObject jsonObject = JSONObject.parseObject(json);
                insertRow(sheet,jsonObject,fields); // 插入行
            }
        }
    }

    /**
     * 填充每行的数据
     * @param sheet
     * @param jsonObject
     * @param fields
     */
    private void insertRow(Sheet sheet,JSONObject jsonObject,String[] fields){
        Row insertRow = sheet.createRow(curRow++);
        int colCount = fields.length;
        for (int i = 0; i < colCount; i++){
            Cell itemCell = insertRow.createCell(i);
            if (jsonObject.containsKey(fields[i])){
                checkTypeAndFillCellValue(itemCell,jsonObject.get(fields[i]));
            }else {
                itemCell.setCellValue("");
            }
        }
    }

    /**
     * 检查json的类型并进行处理
     * 目前日期类型作为字符串进行处理，为进行特殊处理
     * @param cell
     * @param object
     */
    private void checkTypeAndFillCellValue(Cell cell,Object object){
        if (object instanceof Integer || object instanceof Long){
            cell.setCellValue(((Number)object).longValue());
        }else if (object instanceof Boolean){
            cell.setCellValue((Boolean)object);
        }else {
            cell.setCellValue(object.toString());
        }
    }

    /**
     * 根据内容对excel表格进行自动缩放
     * @param sheet
     * @param totalCol
     */
    private void autoSizeColumWidth(Sheet sheet,int totalCol){
        for (int i = 0; i < totalCol; i++){
            sheet.autoSizeColumn(i);
            sheet.setColumnWidth(i,sheet.getColumnWidth(i) * 17/10); // 将默认的列宽放大1.7倍
        }
    }

    /**
     * 将excel保存为文件
     * @param workbook
     * @param fileName
     */
    @SneakyThrows
    private String saveFile(Workbook workbook, String fileName){
        if (Files.notExists(Paths.get(excelPath))){
            Files.createDirectories(Paths.get(excelPath));
        }
        String filePathByName = excelPath + "//" + fileName + ".xlsx";
        try (FileOutputStream outputStream = new FileOutputStream(filePathByName)) {
            workbook.write(outputStream);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return filePathByName;
    }
}
