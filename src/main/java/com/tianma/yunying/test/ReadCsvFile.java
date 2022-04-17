package com.tianma.yunying.test;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.charset.Charset;
import java.util.ArrayList;

import com.csvreader.CsvReader;

/*
 *包名:com.kdzwy.cases
 *作者:Adien_cui
 *时间:2017-9-25  下午4:36:29
 *描述:读取csv文件
 **/
public class ReadCsvFile {
    public static void readCsvFile(String filePath){
        try {
            ArrayList<String[]> csvList = new ArrayList<String[]>(); 
            CsvReader reader = new CsvReader(filePath,',',Charset.forName("GBK"));
//          reader.readHeaders(); //跳过表头,不跳可以注释掉

            while(reader.readRecord()){
                csvList.add(reader.getValues()); //按行读取，并把每一行的数据添加到list集合
            }
            reader.close();
            System.out.println("读取的行数："+csvList.size());

            for(int row=0;row<csvList.size();row++){
                System.out.println("-----------------");
                //打印每一行的数据
                System.out.print(csvList.get(row)[0]+",");
                System.out.print(csvList.get(row)[1]+",");
                System.out.print(csvList.get(row)[2]+",");
                System.out.println(csvList.get(row)[3]+",");
                //如果第一列（即姓名列）包含lisa，则打印出lisa的年龄
                if(csvList.get(row)[0].equals("lisa")){  
                    System.out.println("lisa的年龄为："+csvList.get(row)[2]);
                }
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) {
        String filePath = "D:\\CodePath\\test\\联排计划展示基础表.csv";


        try {
            // 创建CSV读对象
            CsvReader csvReader = new CsvReader(filePath);

            // 读表头
            csvReader.readHeaders();
            while (csvReader.readRecord()){
                // 读一整行
                System.out.println(csvReader.getRawRecord());
                // 读这行的某一列
                System.out.println(csvReader.get("Link"));
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}