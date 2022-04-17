package com.tianma.yunying.test;

import com.tianma.yunying.entity.GanttTask;
import com.tianma.yunying.util.Excel_Util;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

public class RunGantt_Test {
    public static void main(String[] args) throws Exception {
        String filelName = "D:\\CodePath\\test\\副本2022年M+3实验需求Rev 03-整0318.xlsx";
        Excel_Util.workbook = new XSSFWorkbook(filelName);
        DateFormat format=new SimpleDateFormat("yyyy/MM/dd");
        GanttTask parentTask = new GanttTask();
        GanttTask ganttTask_input = new GanttTask();
        GanttTask ganttTask_output = new GanttTask();
        HashMap<String,String[]> Run_sheet = new HashMap<>();
        String sheet_name = "透视表";
        XSSFSheet sheet = Excel_Util.workbook.getSheet("透视表");
        int cur_row=sheet.getLastRowNum();
        String[] table  = {"0","0","0","0","0","0"};
        System.out.println(cur_row);
        for(int i = 4 ; i < cur_row; i++){
            System.out.println(Excel_Util.readExcelData(filelName,sheet_name,i,0));
            table  = new String[]{"0", "0", "0", "0", "0","0"};//数组第一个至第五个为五个厂的对应数量,第六个为统计量
            for(int j = 1; j <=6; j++){
                String tmp_var = Excel_Util.readExcelData(filelName,sheet_name,i,j);
                if(!tmp_var.equals(""))
                    table[j-1] = tmp_var;
            }
            Run_sheet.put(Excel_Util.readExcelData(filelName,sheet_name,i,0),table);
        }

        sheet = Excel_Util.workbook.getSheet("整理");
        int cur_row2=sheet.getLastRowNum();
        for (Map.Entry<String, String[]> entry : Run_sheet.entrySet()) {
            System.out.println("Key = " + entry.getKey() + ", Value = " + entry.getValue());
            for(int i = 1; i < cur_row2;i++){
                String tmp = Excel_Util.readExcelData(filelName,"整理",i,6);
                if(tmp.equals(entry.getKey())){
                    parentTask.setParent(Excel_Util.readExcelData(filelName,"整理",i,1));
                    System.out.println(parentTask.getParent());
                    parentTask.setId(tmp);
                    parentTask.setRender("split");
                    parentTask.setOpen("true");
                    parentTask.setColor("rgba(0,0,0,0)");
                    ganttTask_input.setId(tmp+"计划投入");
                    ganttTask_output.setId(tmp+"计划产出");
                    parentTask.setDesc(Excel_Util.readExcelData(filelName,"整理",i,0));
                    if(!Excel_Util.readExcelData(filelName,"整理",i,7).equals("")){
                        parentTask.setNumber(Integer.parseInt(Excel_Util.readExcelData(filelName,"整理",i,7)));
                    }
                    else parentTask.setNumber(0);
//                    parentTask.setStart_date(Excel_Util.readExcelData(filelName,"整理",i,8));
//                    parentTask.setEnd_date_text(Excel_Util.readExcelData(filelName,"整理",i,9));
                    ganttTask_input.setColor("rgba(255,165,0,0.5)");
                    ganttTask_input.setParent(parentTask.getId());
                    ganttTask_output.setColor("rgba(192,192,192,0.5)");
                    ganttTask_output.setParent(parentTask.getId());
                    String t1 = Excel_Util.readExcelData(filelName,"整理",i,8);
                    String t2 = Excel_Util.readExcelData(filelName,"整理",i,9);
                    if(!t1.equals("")){
                        Date time_date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.valueOf(Excel_Util.readExcelData(filelName,"整理",i,8)));
                        ganttTask_input.setStart_date(format.format(time_date));
                        ganttTask_input.setUse_amount(Excel_Util.readExcelData(filelName,"整理",i,10));
                    }
                    if(!t2.equals("")) {
                        Date time_date2 = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.valueOf(Excel_Util.readExcelData(filelName,"整理",i,9)));
                        ganttTask_output.setStart_date(format.format(time_date2));
                        ganttTask_output.setUse_amount(Excel_Util.readExcelData(filelName,"整理",i,11));
                    }
                    System.out.println(parentTask);
                    System.out.println(ganttTask_input);
                    System.out.println(ganttTask_output);
                }
            }
        }



    }
}
