package com.excel.tool.core.config;

import lombok.AccessLevel;
import lombok.Data;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;

import java.util.Collections;
import java.util.List;

/**
 * excel的导出的配置类
 * 以下是配置项
 */
@Data
@Slf4j
public class ExcelExportConfig<T> {
    private String excelFileName; // 导出的excel名称
    private String sheetName = "sheet1"; // excel工作簿的名称
    private boolean isShowHeader = false; // 是否需要表头
    private boolean isShowSummary = false; // 是否需要汇总
    private boolean isShowQuery = false; // 是否需要查询项

    private String[] titles; // 导出文件的表头
    @Setter(AccessLevel.NONE)
    private List<T> dataSource; // 数据列表

    private String header = ""; // 表头

    public ExcelExportConfig(String excelFileName,String sheetName,boolean isShowHeader,String[] titles){
        this.excelFileName = excelFileName;
        this.isShowHeader = isShowHeader;
        this.titles = titles;
        if (StringUtils.isNotBlank(sheetName)){
            this.sheetName = sheetName;
        }
    }

    /**
     * 填充数据源
     * @param dataList
     */
    public void fillDataSource(List<T> dataList){
        this.dataSource = Collections.unmodifiableList(dataList);
    }
}
