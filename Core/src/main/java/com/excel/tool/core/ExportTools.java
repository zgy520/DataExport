package com.excel.tool.core;

import com.excel.tool.core.config.ExcelExportConfig;
import lombok.SneakyThrows;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

public class ExportTools {

    //@Value("${excel.export.filePath}")
    private String excelPath = "export";
    private int curRow = 0;

    // 导出excel
    public void generateExcel(ExcelExportConfig excelExportConfig) throws Exception {
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

        // 自动缩放
        autoSizeColumWidth(activeSheet,totalColLen);

        // 保存文件
        saveFile(workbook,excelExportConfig.getExcelFileName());
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
        headerCell.setCellValue(header); // 赋值
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
            titleCell.setCellValue(titles[i]);
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
        }
    }

    /**
     * 将excel保存为文件
     * @param workbook
     * @param fileName
     */
    @SneakyThrows
    private void saveFile(Workbook workbook, String fileName){
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
    }
}
