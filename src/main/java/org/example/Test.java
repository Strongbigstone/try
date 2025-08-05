package org.example;

import cn.hutool.core.io.FileUtil;

import java.io.FileInputStream;
import java.io.FileNotFoundException;

public class Test {

    public static void main(String[] args) throws FileNotFoundException {
        FileInputStream in = new FileInputStream("D:\\DK\\Desktop\\测试1.xls");
        FileUtil.writeFromStream(in, "D:\\DK\\Desktop\\测试2.xlsx");
    }
}
