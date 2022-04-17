package com.tianma.yunying.util;

import java.io.*;

public class FileUtil {
//    public static void main(String[] args) throws IOException {
//        FileUtil.write("D:\\1.html"); //运行主方法
//    }
    public static void write(String path,String content)
            throws IOException {
        //将写入转化为流的形式
        BufferedWriter bw = new BufferedWriter(new FileWriter(path));
        //一次写一行
//        String ss = content;
        bw.write(content);
        bw.newLine();  //换行用
        //关闭流
        bw.close();
        System.out.println("写入成功");
    }

    //  // param folderPath 文件夹完整绝对路径
    public static void delFolder(String folderPath) {
        try {
            delAllFile(folderPath); // 删除完里面所有内容
            //不想删除文佳夹隐藏下面
//          String filePath = folderPath;
//          filePath = filePath.toString();
//          java.io.File myFilePath = new java.io.File(filePath);
//          myFilePath.delete(); // 删除空文件夹
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    public static boolean delAllFile(String path) {
        boolean flag = false;
        File file = new File(path);
        if (!file.exists()) {
            return flag;
        }
        if (!file.isDirectory()) {
            return flag;
        }
        String[] tempList = file.list();
        File temp = null;
        for (int i = 0; i < tempList.length; i++) {
            if (path.endsWith(File.separator)) {
                temp = new File(path + tempList[i]);
            } else {
                temp = new File(path + File.separator + tempList[i]);
            }
            if (temp.isFile()) {
                temp.delete();
            }
            if (temp.isDirectory()) {
                delAllFile(path + "/" + tempList[i]);// 先删除文件夹里面的文件
                delFolder(path + "/" + tempList[i]);// 再删除空文件夹
                flag = true;
            }
        }
        return flag;
    }


}
