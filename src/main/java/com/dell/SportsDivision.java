package com.dell;

import com.dell.util.Tool;
import org.junit.Test;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.*;

public class SportsDivision {
    @Test
    public void test1() throws IOException {
        List<String> list = Arrays.asList("李媛媛", "魏子超", "许向南", "王怀", "刘欣程", "俞俊",
                "张敏", "唐伟", "钦程", "李子钰", "张羽铖", "谭钊全", "贺缙殷", "白宏禹", "林祥",
                "郎金刚", "尚含溪", "郭永明", "曾蕊", "陈俊杰", "寇佳军", "郭波", "陈林", "高斌",
                "李远芳", "郑创", "杨锐", "刘宇航", "张永涛", "姚永晴", "潘雯", "杨为茗", "姜佳伟",
                "张磊", "彭然", "汪伟鹏", "付惠玲", "于敏亮", "吴凤岐", "李长茂", "陈亚", "何鑫",
                "官睿", "叶陈锋", "李佳澳", "彭涛", "付文飞", "买春生", "姜雯婷", "张培杰", "冯春",
                "雷职菱", "胡博", "九组替补", "黎俊杰", "任万林", "秦超", "张金辉", "孟贤洁", "徐阳");
        Map<String,String> all = new HashMap<>();
        int flag = 1;
        List<String> temp = new ArrayList<>();
        //所有人存入map
        for (String s : list) {
            temp.add(s);
            if (temp.size()==6){
                for (String s1 : temp) {
                    switch (flag){
                        case 1:
                            all.put(s1,"第一组");
                            break;
                        case 2:
                            all.put(s1,"第二组");
                            break;
                        case 3:
                            all.put(s1,"第三组");
                            break;
                        case 4:
                            all.put(s1,"第四组");
                            break;
                        case 5:
                            all.put(s1,"第五组");
                            break;
                        case 6:
                            all.put(s1,"第六组");
                            break;
                        case 7:
                            all.put(s1,"第七组");
                            break;
                        case 8:
                            all.put(s1,"第八组");
                            break;
                        case 9:
                            all.put(s1,"第九组");
                            break;
                        default:
                            all.put(s1,"第十组");
                            break;
                    }
                }
                temp.clear();
                flag++;
            }
        }
        all.remove("九组替补");
        temp.clear();
        for (Map.Entry<String, String> en : all.entrySet()) {
            temp.add(en.getKey()+":"+en.getValue());
        }//所有人的数据，打乱写入excel
        Collections.shuffle(temp);
        Tool.listInsertExcel(temp);
    }
}
