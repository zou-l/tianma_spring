package com.tianma.yunying.test;

import com.tianma.yunying.entity.GanttTask;
import com.tianma.yunying.util.Excel_Util;
import com.tianma.yunying.util.StringUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;

public class Gantt_Test_new {
    public static void main(String[] args) throws Exception {
        String fileName = "D:\\CodePath\\test\\1_test_new.xlsx";
        String sheetName = "编码规则";
        Excel_Util.workbook = new XSSFWorkbook(fileName);
        int rela_row = Excel_Util.workbook.getSheet(sheetName).getLastRowNum();
        HashMap<String,String> code_Fir= new HashMap<>();
        HashMap<String,String> code_Sec= new HashMap<>();
        HashMap<String,String> code_Thi= new HashMap<>();
        HashMap<String,String> code_Fou= new HashMap<>();
        DateFormat format=new SimpleDateFormat("yyyy/MM/dd");
        String cur_sheet_name = "";
        GanttTask ganttTask = new GanttTask();
        GanttTask parentTask = new GanttTask();
        GanttTask fatherTask = new GanttTask();
        GanttTask tmptask = new GanttTask();
        GanttTask tmptask2 = new GanttTask();
        Calendar cal = Calendar.getInstance();
        String tmp_start_date = "";
        Double tmp_use = 0.0;
        Boolean isStart = true;

        HashMap<Integer, String> sheet_name = new HashMap<>();
        sheet_name.put(1, "ARRAY计划");
        sheet_name.put(2, "EVEN计划");
        sheet_name.put(3, "TPOT计划");
        sheet_name.put(4, "EAC计划");
        sheet_name.put(5, "MODULE计划");

        String tmp_str = "*";
        int tmp_int = 1;
        //将编码规则读入字典
        while(!tmp_str.equals("")&& tmp_int <= rela_row){
            tmp_str = Excel_Util.readExcelData(fileName,sheetName,tmp_int,0);
            System.out.println(tmp_str);
            System.out.println(Excel_Util.readExcelData(fileName,sheetName,tmp_int,1));
            code_Fir.put(Excel_Util.readExcelData(fileName,sheetName,tmp_int,1),Excel_Util.readExcelData(fileName,sheetName,tmp_int,0));
            tmp_int++;
        }
        tmp_int = 1;
        tmp_str = "*";
        while(!tmp_str.equals("")&& tmp_int <= rela_row){
            tmp_str = Excel_Util.readExcelData(fileName,sheetName,tmp_int,2);
            System.out.println(tmp_str);
            System.out.println(Excel_Util.readExcelData(fileName,sheetName,tmp_int,3));
            code_Sec.put(Excel_Util.readExcelData(fileName,sheetName,tmp_int,3),Excel_Util.readExcelData(fileName,sheetName,tmp_int,2));
            tmp_int++;
        }
        System.out.println(code_Sec);
        tmp_int = 1;
        tmp_str = "*";

        while(!tmp_str.equals("")&& tmp_int <= rela_row){
            tmp_str = Excel_Util.readExcelData(fileName,sheetName,tmp_int,4);
            System.out.println(tmp_str);
            System.out.println(Excel_Util.readExcelData(fileName,sheetName,tmp_int,5));
            code_Thi.put(Excel_Util.readExcelData(fileName,sheetName,tmp_int,5),Excel_Util.readExcelData(fileName,sheetName,tmp_int,4));
            tmp_int++;
        }
        tmp_int = 1;
        tmp_str = "*";
        while(!tmp_str.equals("") && tmp_int <= rela_row){
            tmp_str = Excel_Util.readExcelData(fileName,sheetName,tmp_int,6);
            System.out.println(tmp_str);
            System.out.println(Excel_Util.readExcelData(fileName,sheetName,tmp_int,7));
            code_Fou.put(Excel_Util.readExcelData(fileName,sheetName,tmp_int,7),Excel_Util.readExcelData(fileName,sheetName,tmp_int,6));
            tmp_int++;
        }


        for(int sheet_index = 1; sheet_index <=5; sheet_index++){
            //获得当前表的各种信息
            cur_sheet_name = sheet_name.get(sheet_index);
            int cur_row = Excel_Util.readrowNum(fileName,cur_sheet_name);
            int cur_col = Excel_Util.readcolNum(fileName, cur_sheet_name);
            int cur_mark = Excel_Util.readWantCol(fileName,cur_sheet_name,0,"Mark");
            int cur_target = Excel_Util.readWantCol(fileName,cur_sheet_name,0,"投入目的");
            int cur_type = Excel_Util.readWantCol(fileName,cur_sheet_name,0,"IN/OUT");
            System.out.println("!!!!!!!1");
            System.out.println(cur_type);

                //遍历五个工厂
                for(int row = 1; row <= cur_row; row++){
                    parentTask.setColor("rgba(0,0,0,0)");
                    parentTask.setDuration(100);
                    String sub_string = Excel_Util.readExcelData(fileName, cur_sheet_name, row, cur_mark).substring(0,4);
//                    System.out.println(code_Sec.get(sub_string.substring(1,2)));
                    parentTask.setId(code_Fir.get(sub_string.substring(0,1))+code_Sec.get(sub_string.substring(1,2))+code_Thi.get(sub_string.substring(2,3))+code_Fou.get(sub_string.substring(3,4)));
                    parentTask.setText(code_Sec.get(sub_string.substring(1,2))+code_Fou.get(sub_string.substring(3,4))+code_Thi.get(sub_string.substring(2,3)));
                    parentTask.setRender("split");

//
                    fatherTask.setId(code_Fir.get(sub_string.substring(0,1))+code_Sec.get(sub_string.substring(1,2))+code_Thi.get(sub_string.substring(2,3)));
                    fatherTask.setText(code_Fir.get(sub_string.substring(0,1))+code_Sec.get(sub_string.substring(1,2))+code_Thi.get(sub_string.substring(2,3)));
                    fatherTask.setColor("rgba(0,0,0,0)");
//                    fatherTask.setNumber(Integer.parseInt(sheet_relation.get(sub_string).split("\\^")[1]));
                    fatherTask.setOpen("true");
                    parentTask.setParent(fatherTask.getId());
                    parentTask.setStart_date(String.valueOf(format.format(org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.valueOf(Excel_Util.readExcelData(fileName, cur_sheet_name, 0, cur_type+1))))));
                    cal.setTime(format.parse(parentTask.getStart_date()));
                    cal.add(Calendar.DATE,parentTask.getDuration()-1);
                    parentTask.setEnd_date_text(String.valueOf(format.format(cal.getTime())));
//                    System.out.println(parentTask);

                    ganttTask.setParent(parentTask.getId());
                    if(Excel_Util.readExcelData(fileName, cur_sheet_name, row, cur_type).equals("IN")){
                        ganttTask.setId(Excel_Util.readExcelData(fileName, cur_sheet_name, row, cur_mark) + "投入");
                        if (sheet_index == 1) {
                            ganttTask.setColor("rgba(255,165,0,0.5)");
                            parentTask.setGantt_type("ARRAY计划");
                        } else if (sheet_index == 2) {
                            ganttTask.setColor("rgba(255,165,0,0.5)");
                            parentTask.setGantt_type("EVEN计划");
                        } else if (sheet_index == 3) {
                            ganttTask.setColor("rgba(255,165,0,0.5)");
                            parentTask.setGantt_type("TPOT计划");
                        } else if (sheet_index == 4) {
                            ganttTask.setColor("rgba(255,165,0,0.5)");
                            parentTask.setGantt_type("EAC计划");
                        } else if (sheet_index == 5) {
                            ganttTask.setColor("rgba(255,165,0,0.5)");
                            parentTask.setGantt_type("MODULE计划");
                        }

                    } else {
                        ganttTask.setId(Excel_Util.readExcelData(fileName, cur_sheet_name, row, cur_mark) + "产出");
                        ganttTask.setColor("rgba(192,192,192,0.5)");
                    }

                    //创建parent任务
                    if(Excel_Util.readExcelData(fileName, cur_sheet_name, row, cur_type).equals("IN")){
//                    parentTask.setFactory_number(String.valueOf(temp_i));
//                        parentTask.setFactory_number(String.valueOf(sheet_factory.get(parentTask.getText())));
//                        String tmp_str = "";
//                        for(int tmp =0; tmp < temp_i; tmp++)
//                            tmp_str += String.valueOf(temp_i);
//                        parentTask.setNumber(Integer.parseInt(tmp_str));
                        System.out.println(parentTask);
                    }
//                    System.out.println(ganttTask);

                    ganttTask.setDuration(0);
                    ganttTask.setStart_date("");
                    ganttTask.setUse_amount("");
                    isStart = true;
                    int cur_index = 0;
                    for (int col = cur_type+1; col <= cur_col; col++) {
                        String read_res = Excel_Util.readExcelData(fileName, cur_sheet_name, row, col);
                        if (!read_res.equals("") && !read_res.equals("0") && isStart && isStart == true) {
                            String time_Str = Excel_Util.readExcelData(fileName, cur_sheet_name, 0, col);
                            Date time_date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.valueOf(time_Str));
                            ganttTask.setStart_date(format.format(time_date));
                            ganttTask.setDuration(1);
                            ganttTask.setUse_amount(Excel_Util.readExcelData(fileName, cur_sheet_name, row, col));
                            tmp_start_date = ganttTask.getStart_date();
                            tmp_use = Double.parseDouble(ganttTask.getUse_amount());
                            tmptask = ganttTask;
                            tmptask2 = ganttTask;
                            tmptask2.setStart_date(time_Str);
                            isStart = false;
                            cur_index = col;
                        } else if (!read_res.equals("") && !read_res.equals("0") && isStart == false) {
                            tmp_use += Double.parseDouble(read_res);
                            tmptask2.setStart_date(Excel_Util.readExcelData(fileName, cur_sheet_name, 0, col));
                            cur_index = col;
                            if(col == cur_col - 1){
                                String time_Str = Excel_Util.readExcelData(fileName, cur_sheet_name, 0, cur_index);
                                Date time_date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.valueOf(time_Str));
                                tmptask.setEnd_date_text(format.format(time_date));
                            }
                        } else if (col == cur_col - 1) {
                            tmptask.setUse_amount(String.valueOf(tmp_use));
                            String time_Str = tmptask2.getStart_date();
//                            System.out.println(time_Str);
//                            Date time_date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.parseDouble(time_Str));
//                            int duration = (int) ((time_date.getTime() - (format.parse((tmp_start_date)).getTime())) / 86400000);
//                            tmptask.setDuration(tmptask.getDuration() + duration);
                            tmptask.setStart_date(tmp_start_date);
                            String str_time = Excel_Util.readExcelData(fileName, cur_sheet_name, 0, cur_index);
                            if (StringUtils.isNumeric(str_time) && !str_time.equals("")){
                                String time_Str_end = Excel_Util.readExcelData(fileName, cur_sheet_name, 0, cur_index);
                                Date time_date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.valueOf(time_Str_end));
                                tmptask.setEnd_date_text(format.format(time_date));
                            }
                            else
                                tmptask.setEnd_date_text("0");
                            if(!parentTask.getText().equals("")){
                                tmptask.setText(parentTask.getText());
                            }
                            else {
                                tmptask.setText(fatherTask.getText());
                            }
//                        tmptask.setOpen("true");
//                            ganttTaskMapper.insertTask(tmptask);
                            System.out.println(tmptask);
                        }
                    }

                }

        }

    }
}
