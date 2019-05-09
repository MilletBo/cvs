package com.millet;

import com.millet.excelUtil.ExcelUtil;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.io.*;
import java.util.*;

@RunWith(SpringRunner.class)
@SpringBootTest(classes = ExcelUtil.class)
public class CvsApplicationTests {

    @Test
    public void testConstructor1() {
        //用排序的Map且Map的键应与ExcelCell注解的index对应
        Map<String,String> map = new LinkedHashMap<>();
        map.put("a","姓名");
        map.put("b","年龄");
        map.put("c","性别");
        map.put("d","出生日期");
        Collection<Object> dataset=new ArrayList<Object>();
        dataset.add(new Model("", "", "",new Date()));
        dataset.add(new Model(null, null, null,new Date()));
        dataset.add(new Model("王五", "34", "男",new Date()));
        File f=new File("/Users/milletbo/Downloads/test.xls");
        OutputStream out = null;
        try {
            out = new FileOutputStream(f);
            ExcelUtil.exportExcel(map, dataset, out);
            out.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Test
    public void exportXls() throws IOException {
        List<Map<String,Object>> list = new ArrayList<Map<String,Object>>();
        Map<String,Object> map =new LinkedHashMap<String, Object>();
        map.put("name", "米酒");
        map.put("age", 33);
        map.put("birthday", new Date());
        map.put("sex", "女");
        Map<String,Object> map2 =new LinkedHashMap<String, Object>();
        map2.put("name", "测试是否是中文长度不能自动宽度.测试是否是中文长度不能自动宽度.");
        map2.put("age", null);
        map2.put("sex", null);
        map.put("birthday",new Date());
        Map<String,Object> map3 =new LinkedHashMap<String, Object>();
        map3.put("name", "张三");
        map3.put("age", 12);
        map3.put("sex", "男");
        map3.put("birthday",new Date());
        list.add(map);
        list.add(map2);
        list.add(map3);
        // 指定cvs列名
        Map<String,String> map1 = new LinkedHashMap<>();
        map1.put("name","姓名");
        map1.put("age","年龄");
        map1.put("birthday","出生日期");
        map1.put("sex","性别");
        File f = new File("/Users/milletbo/Downloads/test1.xls");
        OutputStream out = new FileOutputStream(f);
        ExcelUtil.exportExcel(map1,list, out );
        out.close();
    }


}
