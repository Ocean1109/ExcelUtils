import com.ocean.excelutils.ExcelUtils;
import org.junit.Test;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * @author huhaiyang
 * @date 2022/9/27
 */
public class ExcelUtilsTest {
    @Test
    public void testReadFromXls(){
        try {
            List<List<String>> res = ExcelUtils.readFromFile(new File("src/main/java/com/ocean/excelutils/workbook.xls"));
            res.stream().forEach(i->{
                i.stream().forEach(j->{
                    System.out.print(j+" ");
                });
                System.out.println();
            });
            List<List<String>> res2 = ExcelUtils.readFromFile(new File("src/main/java/com/ocean/excelutils/workbook.xlsx"));
            res2.stream().forEach(i->{
                i.stream().forEach(j->{
                    System.out.print(j+" ");
                });
                System.out.println();
            });
        }catch (Exception e){
            e.printStackTrace();
        }

    }

    @Test
    public void testWriteToXls(){
        try {
            List<String> attributes = Arrays.asList("属性10","属性11");
            List<List<String>> data = Arrays.asList(Arrays.asList("1","2"),Arrays.asList("3","4"));
            ExcelUtils.writeToFile(attributes,data,new File("src/main/java/com/ocean/excelutils/write.xls"));
            ExcelUtils.writeToFile(attributes,data,new File("src/main/java/com/ocean/excelutils/write.xlsx"));
        }catch (Exception e){
            e.printStackTrace();
        }
    }
}
