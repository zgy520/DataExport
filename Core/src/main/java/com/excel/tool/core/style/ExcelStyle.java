package com.excel.tool.core.style;

import org.apache.poi.ss.usermodel.*;

/**
 * 导出的各类样式信息
 */
public class ExcelStyle {

    /**
     * 获取excel表头的样式
     * @param sheet
     * @return
     */
    public static CellStyle getHeaderCellStyle(Sheet sheet){
        Font headerFont =sheet.getWorkbook().createFont();
        headerFont.setFontHeightInPoints((short)16);
        headerFont.setBold(true);
        headerFont.setColor(IndexedColors.BLACK.index);
        CellStyle headerStyle =sheet.getWorkbook().createCellStyle();
        headerStyle.setAlignment(HorizontalAlignment.GENERAL); // 居中
        headerStyle.setFont(headerFont);
        return headerStyle;
    }

    /**
     * 获取表头的样式
     * @param sheet
     * @return
     */
    public static CellStyle getTitleCellStyle(Sheet sheet){
        Font titleFont = sheet.getWorkbook().createFont();
        titleFont.setFontHeightInPoints((short)14);
        titleFont.setBold(true);
        CellStyle titleStyle = sheet.getWorkbook().createCellStyle();
        titleStyle.setFont(titleFont);
        titleStyle.setAlignment(HorizontalAlignment.CENTER);
        return titleStyle;
    }
}
