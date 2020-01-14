import com.excel.tool.core.ExportTools;
import com.excel.tool.core.config.ExcelExportConfig;
import com.excel.tool.core.model.User;
import org.junit.Before;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.mockito.Spy;
import org.springframework.test.context.event.annotation.BeforeTestMethod;
import org.springframework.test.context.junit4.SpringRunner;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

@RunWith(SpringRunner.class)
public class ExcelToolTest {
    @Spy
    private List<User> userList = new ArrayList<>();
    @Before
    public void fillUserList(){
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
        try {
            Date birthOne = sdf.parse("1991-02-14");
            Date birthTwo = sdf.parse("2015-12-05");
            User user = new User("zgy","1991-02-14",28,"河北衡水");
            userList.add(user);
            User userTwo = new User("张广义","2015-12-05",23,"东疆湾");
            userList.add(userTwo);
        }catch (ParseException ex){

        }

    }
    @Test
    public void testGenerateExcelMethod(){
        ExportTools exportTools = new ExportTools();

        String[] titles = {"姓名","初始日期","年龄","地址"};
        String[] fields = {"name","birthDate","age","address"};
        ExcelExportConfig<User> excelConfig = new ExcelExportConfig<User>("测试导出文件","基本信息",false,titles,fields);
        excelConfig.setHeader("这是测试的表头");
        excelConfig.fillDataSource(userList);
        try {
            exportTools.generateExcel(excelConfig);
        }catch (Exception ex){
            ex.printStackTrace();
        }

    }
}
