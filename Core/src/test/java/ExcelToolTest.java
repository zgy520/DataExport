import com.excel.tool.core.ExportTools;
import com.excel.tool.core.config.ExcelExportConfig;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.test.context.junit4.SpringRunner;

@RunWith(SpringRunner.class)
public class ExcelToolTest {
    @Test
    public void testGenerateExcelMethod(){
        ExportTools exportTools = new ExportTools();
        String[] titles = {"姓名","年龄"};
        ExcelExportConfig<String> excelConfig = new ExcelExportConfig<>("测试导出文件","基本信息",true,titles);
        excelConfig.setHeader("这是测试的表头");
        try {
            exportTools.generateExcel(excelConfig);
        }catch (Exception ex){
            ex.printStackTrace();
        }

    }
}
