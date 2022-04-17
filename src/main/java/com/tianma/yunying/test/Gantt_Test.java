package com.tianma.yunying.test;

import com.tianma.yunying.entity.GanttTask;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;

public class Gantt_Test {
    public static void main(String[] args) throws Exception {
        long start = System.currentTimeMillis();
//        String filelName = "D:\\CodePath\\test\\联排编码逻辑.xlsx";
        String filelName = "D:\\CodePath\\test\\1_test.xlsx";
        PoiExcelTest.workbook = new XSSFWorkbook(filelName);
        HashMap<Integer, String> sheet_name = new HashMap<>();
        HashMap<String,String> sheet_parent = new HashMap<>();
        HashMap<String,String> sheet_relation = new HashMap<>();
        int rela_col = PoiExcelTest.readcolNum(filelName,"对应顺序表");
        int rela_row = PoiExcelTest.readrowNum(filelName,"对应顺序表");
        for(int i = 1; i <= rela_row; i++){
            for(int j = 0; j < rela_col; j++){
//                System.out.println(PoiExcelTest.readExcelData(filelName,"对应顺序表",i,j));
                String tmp = PoiExcelTest.readExcelData(filelName,"对应顺序表",i,2)+"^"+PoiExcelTest.readExcelData(filelName,"对应顺序表",i,0);
                sheet_relation.put(PoiExcelTest.readExcelData(filelName,"对应顺序表",i,1),tmp);
            }
        }
        System.out.println(sheet_relation);


        sheet_name.put(1, "ARRAY计划");
        sheet_name.put(2, "EVEN计划");
        sheet_name.put(3, "TPOT计划");
        sheet_name.put(4, "EAC计划");
        sheet_name.put(5, "MODULE计划");


        String cur_sheet_name = "";
        int cur_row = 0;
        int cur_col = 0;
        int cur_type=0;
        int cur_mark=0;
        int cur_target=0;
        DateFormat format=new SimpleDateFormat("yyyy/MM/dd");
        GanttTask ganttTask = new GanttTask();
        GanttTask parentTask = new GanttTask();
        GanttTask tmptask = new GanttTask();
        GanttTask tmptask2 = new GanttTask();
        String tmp_start_date = "";
        Double tmp_use = 0.0;
        ganttTask.setOpen("true");
        Boolean isStart = true;
        for (int i = 1; i <= 5; i++) {
            cur_sheet_name = sheet_name.get(i);
            System.out.println(cur_sheet_name);
            cur_col = PoiExcelTest.readcolNum(filelName, cur_sheet_name);

            System.out.println("cur_col:" + cur_col);
            for (int j = 0; j <= cur_col; j++) {
                String tmp = PoiExcelTest.readExcelData(filelName, cur_sheet_name, 0, j);
                if (tmp.equals("Mark")) {
                    cur_mark = j;
                    cur_row = PoiExcelTest.readrowNum2(filelName, cur_sheet_name, cur_mark);
                }
                if(tmp.equals("投入目的")){
                    cur_target = j;
                }
                if (tmp.equals("IN/OUT")) {
                    cur_type = j;
//                    System.out.println(cur_type);
                    break;
                }
            }
            System.out.println("cur_row:" + cur_row);
            Calendar cal = Calendar.getInstance();
            for (int row = 1; row <= cur_row; row++) {
                String var_temp = PoiExcelTest.readExcelData(filelName, cur_sheet_name, row, cur_mark);
                parentTask.setId(var_temp);
                parentTask.setDuration(100);
//                if(PoiExcelTest.readExcelData(filelName, cur_sheet_name, row, cur_target).contains("项目")){
                String sub_string = PoiExcelTest.readExcelData(filelName, cur_sheet_name, row, cur_mark).substring(0,3);
                parentTask.setParent(sheet_relation.get(sub_string).split("\\^")[0]);
                parentTask.setText(PoiExcelTest.readExcelData(filelName, cur_sheet_name, row, cur_target));
                System.out.println(sub_string);
                System.out.println("YYYYYYYYYYYYYYYYYYYY");
                System.out.println(sheet_relation.get(sub_string).split("\\^")[1]);

//                    parentTask.setParent("送样");
//                }
//                else {
//                    parentTask.setParent("pilot_2");
//                }

                parentTask.setStart_date(String.valueOf(format.format(org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.valueOf(PoiExcelTest.readExcelData(filelName, cur_sheet_name, 0, cur_type+1))))));
                ganttTask.setParent(parentTask.getId());
                cal.setTime(format.parse(parentTask.getStart_date().toString()));
                cal.add(Calendar.DATE,parentTask.getDuration());
                parentTask.setEnd_date_text(String.valueOf(format.format(cal.getTime())));
                if (PoiExcelTest.readExcelData(filelName, cur_sheet_name, row, cur_type).equals("IN")) {
                    ganttTask.setId(var_temp + "投入");
                    if (i == 1) {
                        ganttTask.setColor("rgba(255,165,0,0.4)");
                        parentTask.setGantt_type("ARRAY计划");
                        System.out.println(parentTask);
                    } else if (i == 2) {
                        ganttTask.setColor("rgba(255,105,180,0.4)");
                        parentTask.setGantt_type("EVEN计划");
                        System.out.println(parentTask);
                    } else if (i == 3) {
                        ganttTask.setColor("rgba(0,255,0,0.4)");
                        parentTask.setGantt_type("TPOT计划");
                        System.out.println(parentTask);
                    } else if (i == 4) {
                        ganttTask.setColor("rgba(153,50,204,0.4)");
                        parentTask.setGantt_type("EAC计划");
                        System.out.println(parentTask);
                    } else if (i == 5) {
                        ganttTask.setColor("rgba(210,180,140,0.4)");
                        parentTask.setGantt_type("MODULE计划");
                        System.out.println(parentTask);
                    }

                } else {
                    ganttTask.setId(PoiExcelTest.readExcelData(filelName, cur_sheet_name, row, cur_mark) + "产出");
                    ganttTask.setColor("rgba(0,128,128,0.4)");
                }
                ganttTask.setDuration(0);
                ganttTask.setStart_date("");
                ganttTask.setUse_amount("");
                isStart = true;
                for (int col = cur_type+1; col <= cur_col; col++) {
//                    System.out.println(col);
                    String read_res = PoiExcelTest.readExcelData(filelName, cur_sheet_name, row, col);
                    if (!read_res.equals("") && !read_res.equals("0") && isStart && isStart == true) {
                        System.out.println("11111111111111111  "+read_res);
                        String time_Str = PoiExcelTest.readExcelData(filelName, cur_sheet_name, 0, col);
                        Date time_date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.valueOf(time_Str));
                        ganttTask.setStart_date(format.format(time_date));
                        ganttTask.setDuration(1);
                        ganttTask.setUse_amount(PoiExcelTest.readExcelData(filelName, cur_sheet_name, row, col));
                        tmp_start_date = ganttTask.getStart_date();
                        tmp_use = Double.parseDouble(ganttTask.getUse_amount());
                        tmptask = ganttTask;
                        tmptask2 = ganttTask;
                        tmptask2.setStart_date(time_Str);
                        isStart = false;
                    } else if (!read_res.equals("") && !read_res.equals("0") && isStart == false) {
                        tmp_use += Double.parseDouble(read_res);
                        tmptask2.setStart_date(PoiExcelTest.readExcelData(filelName, cur_sheet_name, 0, col));
                    } else if (col == cur_col - 1) {
                        tmptask.setUse_amount(String.valueOf(tmp_use));
                        String time_Str = tmptask2.getStart_date();
//                            System.out.println(time_Str);
                        Date time_date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.parseDouble(time_Str));
                        int duration = (int) ((time_date.getTime() - (format.parse((tmp_start_date)).getTime())) / 86400000);
                        tmptask.setDuration(tmptask.getDuration() + duration);
                        tmptask.setStart_date(tmp_start_date);
                        cal.setTime(format.parse(tmptask.getStart_date()));
                        cal.add(Calendar.DATE,tmptask.getDuration());
                        tmptask.setEnd_date_text(String.valueOf(format.format(cal.getTime())));
                        System.out.println(tmptask);

//                            ganttTaskMapper.insertTask(tmptask);
                    }
                }
//                System.out.println(ganttTask);
            }
        }
        long end = System.currentTimeMillis();
        System.out.println("程序运行时间："+(end-start)+"ms");
        PoiExcelTest.closeExcel();
    }
}