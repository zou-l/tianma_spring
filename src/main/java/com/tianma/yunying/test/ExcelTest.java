package com.tianma.yunying.test;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import com.tianma.yunying.util.ExcelUtil;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.junit.jupiter.api.Test;
import sun.util.calendar.BaseCalendar;

import javax.xml.crypto.Data;
import java.io.File;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

public class ExcelTest {
    private static ActiveXComponent xl = null; //Excel对象(防止打开多个)
    private static Dispatch workbooks = null;  //工作簿对象
    private static Dispatch workbook = null; //具体工作簿
    private static Dispatch sheets = null;// 获得sheets集合对象
    private static Dispatch currentSheet = null;// 当前sheet
    @Test
    public static void main(String[] args) {
        String filePath = "D:\\CodePath\\test\\mail\\1_test.xlsx";
        //        ActiveXComponent axOutlook = new ActiveXComponent("Excel.Application");
        OpenExcel(filePath,true);
        SaveAs("D:\\CodePath\\test\\mail\\1_test.htm");
        releaseSource();

    }

    public static void OpenExcel(String filepath, boolean visible) {
        try {
            initComponents(); //清空原始变量
            ComThread.InitSTA();
            if(xl==null)
                xl = new ActiveXComponent("Excel.Application"); //Excel对象
            xl.setProperty("Visible", new Variant(visible));//设置是否显示打开excel
            if(workbooks==null)
                workbooks = xl.getProperty("Workbooks").toDispatch(); //打开具体工作簿
            workbook = Dispatch.invoke(workbooks, "Open", Dispatch.Method,
                    new Object[] { filepath,
                            new Variant(false), // 是否以只读方式打开
                            new Variant(true),
                            "1",
                            "pwd" },   //输入密码"pwd",若有密码则进行匹配，无则直接打开
                    new int[1]).toDispatch();
        } catch (Exception e) {
            e.printStackTrace();
            releaseSource();
        }
    }
    public static void SaveAs(String filePath){
        Dispatch.call(workbook, "SaveAs",filePath,44);
    }

    private static void initComponents(){
        workbook = null;
        currentSheet = null;
        sheets = null;
    }

    public static void releaseSource(){
        if(xl!=null){
            xl.invoke("Quit", new Variant[] {});
            xl = null;
        }
        workbooks = null;
        ComThread.Release();
        System.gc();
    }
}
