package com.tianma.yunying.test;

import com.tianma.yunying.entity.GanttTask;
import com.tianma.yunying.entity.Gantt_Detail;
import com.tianma.yunying.entity.Gantt_Info;
import com.tianma.yunying.mapper.GanttMapper;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.*;

public class PoiExcelTest {

    @Autowired
    GanttMapper ganttMapper;
    public static XSSFWorkbook workbook; // 工作簿
    public static XSSFSheet sheet; // 工作表
    public static XSSFRow row; // 行
    public static XSSFCell cell; // 列
//    public static void main(String[] args) throws Exception{
//        String filelName = "D:\\CodePath\\test\\联排计划展示基础表 (1).xlsx";
//        HashMap<Integer,String> sheet_name = new HashMap<>();
//        sheet_name.put(1,"ARRAY厂计划");
//        sheet_name.put(2,"EVEN厂计划");
//        sheet_name.put(3,"TPOT厂计划 ");
//        sheet_name.put(4,"EAC厂计划");
//        sheet_name.put(5,"MODULE厂计划");
//        String cur_sheet_name = "";
//        Gantt_Info gantt_info = new Gantt_Info();
//        Gantt_Detail gantt_detail = new Gantt_Detail();
//        int cur_row = 0;
//        int cur_col = 0;
////        for(int i = 1; i <= 5; i++){
////            cur_sheet_name = sheet_name.get(i);
////            if(i == 2)
////                tmp_row = 9;
////            else if(i == 4)
////                tmp_row = 11;
////            for(int row = 1; row <tmp_row; row++){
////                for(int col = 0; col<7; col++){
////                    if(col == 0){
////                        gantt_info.setFactory_type(readExcelData(fielName,cur_sheet_name,row,col));
////                    }
////                    else if(col == 1){
////                        gantt_info.setLabel(readExcelData(fielName,cur_sheet_name,row,col));
////                    }
////                    else if(col == 2){
////                        gantt_info.setDepartment(readExcelData(fielName,cur_sheet_name,row,col));
////                    }
////                    else if(col == 3){
////                        gantt_info.setCustomer(readExcelData(fielName,cur_sheet_name,row,col));
////                    }
////                    else if(col == 4){
////                        gantt_info.setOutput_no(readExcelData(fielName,cur_sheet_name,row,col));
////                    }
////                    else if(col == 5){
////                        gantt_info.setTotal(Double.valueOf(readExcelData(fielName,cur_sheet_name,row,col)).intValue());
////                    }
////                    else if(col == 6){
////                        gantt_info.setIN_OUTPUT(readExcelData(fielName,cur_sheet_name,row,col));
////                    }
////                }
////            }
////        }
//        DateFormat format=new SimpleDateFormat("yyyy/MM/dd");
//        for(int i = 1; i <=5; i++){
//            cur_sheet_name = sheet_name.get(i);
//            cur_row = PoiExcelTest.readrowNum(filelName,cur_sheet_name);
//            cur_col = PoiExcelTest.readcolNum(filelName,cur_sheet_name);
//            for(int row = 1; row < cur_row; row++){
//                gantt_detail.setLabel(PoiExcelTest.readExcelData(filelName,cur_sheet_name,row,1));
//                gantt_detail.setIN_OUTPUT(PoiExcelTest.readExcelData(filelName,cur_sheet_name,row,6));
//                for(int col = 7; col <cur_col; col++){
////                    gantt_detail.setUse_time(PoiExcelTest.readExcelData(filelName,cur_sheet_name,0,col));
//                    String time_Str = PoiExcelTest.readExcelData(filelName,cur_sheet_name,0,col);
//                    Date time_date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.valueOf(time_Str));
//                    gantt_detail.setUse_time(format.format(time_date));
//                    if(PoiExcelTest.readExcelData(filelName,cur_sheet_name,row,col).equals(""))
//                        gantt_detail.setUse_amount("0");
//                    else
//                        gantt_detail.setUse_amount(PoiExcelTest.readExcelData(filelName,cur_sheet_name,row,col));
////                    System.out.println(gantt_detail);
//                }
//            }
//        }
//        SimpleDateFormat sdf = new SimpleDateFormat();// 格式化时间
//        sdf.applyPattern("yyyy-MM-dd HH:mm:ss a");// a为am/pm的标记
//        Date date = new Date();// 获取当前时间
//        System.out.println("现在时间：" + sdf.format(date)); // 输出已经格式化的现在时间（24小时制）
//    }

//    public static void main(String[] args) throws Exception {
//        String filelName = "D:\\CodePath\\test\\联排计划展示基础表 (1).xlsx";
//        String sheetname = "标签关系表";
//        Boolean isSame = true;
//        HashMap<String, List<String>> relation = new HashMap<>();
//        List<Integer> list_row = new LinkedList<>();
//        List<String> list_string = new ArrayList<>();
//        DateFormat format=new SimpleDateFormat("yyyy-MM-dd");
//        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
//        Date time_date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.valueOf(PoiExcelTest.readExcelData(filelName,"ARRAY厂计划",0,7)));
//        int r_row = PoiExcelTest.readrowNum(filelName,sheetname);
//        GanttTask parent_task = new GanttTask();
//        parent_task.setDuration(100);
//        parent_task.setStart_date(format.format(time_date));
//        parent_task.setColor("rgba(0,0,0,0.0)");
//        parent_task.setOpen("true");
//        parent_task.setParent("产品1");
//        parent_task.setRender("split");
////        int col = readcolNum(filelName,sheetname);
//        for(int i = 1; i <= r_row-1; i++){
//            if(PoiExcelTest.readExcelData(filelName,sheetname,i,5).equals("是")){
//                String key = PoiExcelTest.readExcelData(filelName,sheetname,i,7) +"\\*"+PoiExcelTest.readExcelData(filelName,sheetname,i,6);
//                List<String> label = new LinkedList<>();
//                for(int j = 0; j < 5; j++){
//                    String tmp_string = PoiExcelTest.readExcelData(filelName,sheetname,i,j);
//                    if(!list_string.contains(tmp_string)){
//                        list_string.add(tmp_string);
//                        label.add(PoiExcelTest.readExcelData(filelName,sheetname,i,j));
//                    }
//                }
//                System.out.println(label);
//                relation.put(key,label);
//            }
//        }
////        int tmp_row = 0;
////        int tmp_col = 0;
////        for(int i = 1; i <= r_row; i++){
////            if(PoiExcelTest.readExcelData(filelName,sheetname,i,5).equals("是")) {
////                isSame = true;
////                for(int j = i; j < r_row && isSame; j++){
////                    if(!PoiExcelTest.readExcelData(filelName,sheetname,i,0).equals(PoiExcelTest.readExcelData(filelName,sheetname,i+1,0))){
////                        isSame = false;
////                        tmp_row = j;
////                        i = j;
//////                        System.out.println(tmp_row);
////                        list_row.add(tmp_row);
////                        if(j+1==r_row){
////                            j++;
////                            list_row.add(j++);
//////                            System.out.println(j++);
////                        }
////                        break;
////                    }
////
////                }
////            }
////        }
////        int tmp_index = 1;
////        isSame = true;
////        for(int i = 0; i <list_row.size();i++){
////            System.out.println("start_row"+String.valueOf(tmp_index));
////            System.out.println("end_row"+String.valueOf(list_row.get(i)));
////            for(int row = tmp_index; row <= list_row.get(i);row++){
////                for(int col = 0; col < 5 && isSame; col++){
////                    if(!PoiExcelTest.readExcelData(filelName,sheetname,row,col).equals(PoiExcelTest.readExcelData(filelName,sheetname,list_row.get(i),col))){
////
////                    }
////                }
////            }
////            tmp_index =list_row.get(i) + 1;
////        }
////        System.out.println(list_row);
//        System.out.println(relation);
//
////        int tmp_var = 0;
////        for(Map.Entry<String, List<String>> entry : relation.entrySet()){
////            Iterator<String> it = entry.getValue().iterator();
////            while(it.hasNext()){//判断是否有迭代元素
//////                System.out.println(it.next());//输出迭代出的元素
//////                System.out.println("33333333333");
////                String a = it.next();
//////                System.out.println("111111111"+a+"2222222222");
////                if(a.equals("")){
////                    break;
////                }
////                if(tmp_var == 0)
////                    parent_task.setGantt_type("Array计划");
////                else if(tmp_var == 1)
////                    parent_task.setGantt_type("EVEN计划");
////                else if(tmp_var == 2)
////                    parent_task.setGantt_type("TPOT计划");
////                else if(tmp_var == 3)
////                    parent_task.setGantt_type("EAC计划");
////                else if(tmp_var == 4)
////                    parent_task.setGantt_type("MODULE计划");
////                parent_task.setText(entry.getKey().split("\\*")[0].replace("\\",""));
////                parent_task.setDesc(entry.getKey().split("\\*")[1]);
////                parent_task.setId(a);
////                System.out.println(parent_task);
//////                ganttTaskMapper.insertTask(parent_task);
////                tmp_var += 1;
////            }
////            tmp_var = 0;
////            System.out.println("---------------------");
////        }
//
//
//        int tmp_var = 0;
//        int tmp_count = 1;
//        int tmp_full_var = 1;
//        int tmp_factory_number = 1;
//        Boolean is_test = false;
//        Boolean is_Full = false;
//        for(Map.Entry<String, List<String>> entry : relation.entrySet()){
//            String tmp_number = "";
//
////            if(tmp_full_var == 0 && entry.getValue().size() == 5){
////                is_Full = false;
////                parent_task.setFactory_number(String.valueOf(tmp_factory_number));
////            }
////            else if(!is_Full && entry.getValue().size() <5){
////                parent_task.setFactory_number(String.valueOf(tmp_factory_number));
////            }
////            else if(entry.getValue().size() == 5){
////                tmp_factory_number += 1;
////                parent_task.setFactory_number(String.valueOf(tmp_factory_number));
////                is_Full = true;
////            }
//
//
////            if(entry.getValue().size() < 5 && !is_Full){
////                parent_task.setFactory_number(String.valueOf(tmp_factory_number));
////                is_test = true;
////            }
////            else if(entry.getValue().size() < 5 && is_Full){
////                is_Full = false;
////                tmp_factory_number += 1;
////                parent_task.setFactory_number(String.valueOf(tmp_factory_number));
////            }
////            else if(entry.getValue().size() == 5 && tmp_full_var == 1 && !is_Full){
////                tmp_factory_number += 1;
////                parent_task.setFactory_number(String.valueOf(tmp_factory_number));
//////                tmp_factory_number += 1;
////            }
////            else if(entry.getValue().size() == 5 && tmp_full_var == 1 && is_Full){
////                tmp_factory_number += 1;
////                parent_task.setFactory_number(String.valueOf(tmp_factory_number));
////                is_Full = false;
//////                tmp_factory_number += 1;
////            }
////            else if(entry.getValue().size() == 5 && tmp_full_var != 1 && !is_Full){
////                if(is_test){
////                    tmp_factory_number += 1;
////                }
////                parent_task.setFactory_number(String.valueOf(tmp_factory_number));
//////                tmp_factory_number += 1;
////                tmp_full_var = 0;
////                is_Full = true;
////            }
////            else if(entry.getValue().size() == 5 && tmp_full_var != 1 && is_Full){
////                tmp_factory_number += 1;
////                parent_task.setFactory_number(String.valueOf(tmp_factory_number));
////                is_Full = false;
////            }
////            tmp_full_var += 1;
//
//            if(entry.getValue().size() < 5 && !is_Full ){
//                parent_task.setFactory_number(String.valueOf(tmp_factory_number));
//            }
//            else if(entry.getValue().size() == 5 && !is_Full&& !is_test){
//                parent_task.setFactory_number(String.valueOf(tmp_factory_number));
//                is_Full = true;
//            }
//
//            else if(entry.getValue().size() == 5 && is_Full){
//                tmp_factory_number += 1;
//                is_Full = false;
//                is_test = true;
//
//                parent_task.setFactory_number(String.valueOf(tmp_factory_number));
//            }
//            else if(entry.getValue().size() == 5 && is_test){
//                is_test = false;
//                tmp_factory_number += 1;
//                parent_task.setFactory_number(String.valueOf(tmp_factory_number));
//            }
//
////            tmp_factory_number += 1;
//            if(entry.getValue().size() == 5){
//            Iterator<String> it = entry.getValue().iterator();
//            while(it.hasNext()) {//判断是否有迭代元素
//                String a = it.next();
//                if (a.equals("")) {
//                    break;
//                }
//                if (tmp_var == 0) {
//                    parent_task.setGantt_type("Array计划");
//                    tmp_number += String.valueOf(tmp_count);
//                    parent_task.setNumber(Integer.parseInt(tmp_number));
//                } else if (tmp_var == 1) {
//                    parent_task.setGantt_type("EVEN计划");
//                    tmp_number += String.valueOf(tmp_count);
//                    parent_task.setNumber(Integer.parseInt(tmp_number));
//                } else if (tmp_var == 2) {
//                    parent_task.setGantt_type("TPOT计划");
//                    tmp_number += String.valueOf(tmp_count);
//                    parent_task.setNumber(Integer.parseInt(tmp_number));
//                } else if (tmp_var == 3) {
//                    parent_task.setGantt_type("EAC计划");
//                    tmp_number += String.valueOf(tmp_count);
//                    parent_task.setNumber(Integer.parseInt(tmp_number));
//                } else if (tmp_var == 4) {
//                    parent_task.setGantt_type("MODULE计划");
//                    tmp_number += String.valueOf(tmp_count);
//                    parent_task.setNumber(Integer.parseInt(tmp_number));
//                }
//
//                parent_task.setText(entry.getKey().split("\\*")[0].replace("\\", ""));
//                parent_task.setDesc(entry.getKey().split("\\*")[1]);
//                parent_task.setId(a);
//                System.out.println(parent_task);
////                ganttTaskMapper.insertTask(parent_task);
//                tmp_var += 1;
//            }
//            }
//            else {
//                tmp_number = "";
//                Iterator<String> it = entry.getValue().iterator();
//                tmp_var = 5 - entry.getValue().size();
//                System.out.println(tmp_var);
//                Boolean flag = true;
//                int count = 1;
//                List<Integer> list_a = new ArrayList<>();
//                while(it.hasNext()){
//                    String a = it.next();
//                    parent_task.setText(entry.getKey().split("\\*")[0].replace("\\", ""));
//                    parent_task.setDesc(entry.getKey().split("\\*")[1]);
//                    parent_task.setId(a);
//                        if(flag) {
//                            for (int i = 0; i <= tmp_var; i++) {
//                                tmp_number += String.valueOf(tmp_count);
//                            }
//                            for (int i = 1; i <= entry.getValue().size(); i++) {
//                                list_a.add(Integer.parseInt(tmp_number) + i);
//                                System.out.println(Integer.parseInt(tmp_number) + i);
//                            }
//                            flag = false;
//                        }
//                        if(entry.getValue().size() - count == 0){
//                            parent_task.setGantt_type("MODULE计划");
//                        }
//                        if(entry.getValue().size() - count == 1){
//                            parent_task.setGantt_type("EAC计划");
//                        }
//                        if(entry.getValue().size() - count == 2){
//                            parent_task.setGantt_type("TPOT计划");
//                        }
//                        if(entry.getValue().size() - count == 3){
//                            parent_task.setGantt_type("EVEN计划");
//                        }
//                        parent_task.setNumber(list_a.get(count-1));
//                        System.out.println(parent_task);
//                        count++;
////                    }
//                }
//
//            }
//
//            tmp_var = 0;
//            tmp_number ="";
//            tmp_count += 1;
//
//            System.out.println("---------------------");
//        }
//
//
////        HashMap<Integer,String> sheet_name = new HashMap<>();
////        sheet_name.put(1,"ARRAY厂计划");
////        sheet_name.put(2,"EVEN厂计划");
////        sheet_name.put(3,"TPOT厂计划 ");
////        sheet_name.put(4,"EAC厂计划");
////        sheet_name.put(5,"MODULE厂计划");
////        String cur_sheet_name = "";
////        int cur_row = 0;
////        int cur_col = 0;
////        GanttTask ganttTask = new GanttTask();
////        GanttTask tmptask = new GanttTask();
////        GanttTask tmptask2 = new GanttTask();
////        String tmp_start_date = "";
////        Double tmp_use = 0.0;
////        ganttTask.setOpen("true");
////        Boolean isStart = true;
////        for(int i = 1; i<=5; i++){
////            cur_sheet_name = sheet_name.get(i);
////            cur_row = PoiExcelTest.readrowNum(filelName,cur_sheet_name);
////            cur_col = PoiExcelTest.readcolNum(filelName,cur_sheet_name);
////            for(int row = 1; row <=cur_row; row++){
////
////                if(PoiExcelTest.readExcelData(filelName,cur_sheet_name,row,6).equals("计划投入")) {
////                    ganttTask.setId(PoiExcelTest.readExcelData(filelName, cur_sheet_name, row, 1) + "计划投入");
////                    if (i == 1)
////                        ganttTask.setColor("rgba(255,165,0,0.4)");
////                    else if (i == 2)
////                        ganttTask.setColor("rgba(255,105,180,0.4)");
////                    else if (i == 3)
////                        ganttTask.setColor("rgba(0,255,0,0.4)");
////                    else if (i == 4)
////                        ganttTask.setColor("rgba(153,50,204,0.4)");
////                    else if (i == 5)
////                        ganttTask.setColor("rgba(210,180,140,0.4)");
////                }
////                else{
////                    ganttTask.setId(PoiExcelTest.readExcelData(filelName,cur_sheet_name,row,1)+"计划产出");
////                    ganttTask.setColor("rgba(0,128,128,0.4)");
////                }
////                ganttTask.setGantt_type(cur_sheet_name.split("厂")[0]+"计划");
////                ganttTask.setText("test");
////                ganttTask.setDuration(0);
////                ganttTask.setProject_name("project_name");
////                ganttTask.setParent(PoiExcelTest.readExcelData(filelName,cur_sheet_name,row,1));
////                isStart = true;
////                for(int col = 0; col < cur_col; col++){
////                    if(col > 6){
////                        String read_res =  PoiExcelTest.readExcelData(filelName,cur_sheet_name,row,col);
////                        if(!read_res.equals("") &&!read_res.equals("0") && isStart && isStart == true)
////                        {
////                            String time_Str = PoiExcelTest.readExcelData(filelName,cur_sheet_name,0,col);
////                            time_date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.valueOf(time_Str));
////                            ganttTask.setStart_date(format.format(time_date));
////                            ganttTask.setDuration(1);
////                            ganttTask.setUse_amount(PoiExcelTest.readExcelData(filelName,cur_sheet_name,row,col));
////                            tmp_start_date = ganttTask.getStart_date();
////                            tmp_use = Double.parseDouble(ganttTask.getUse_amount());
////                            tmptask = ganttTask;
////                            tmptask2 = ganttTask;
////                            tmptask2.setStart_date(time_Str);
////                            isStart = false;
////                        }
//////                        else if(!read_res.equals("") &&!read_res.equals("0") && isStart == false){
////////                            System.out.println(read_res);
//////                            String time_Str = PoiExcelTest.readExcelData(filelName,cur_sheet_name,0,col);
//////                            Date time_date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.valueOf(time_Str));
//////                            int duration = (int) ((time_date.getTime() - (simpleDateFormat.parse((tmp_start_date)).getTime()))/86400000);
//////                            if(PoiExcelTest.readExcelData(filelName,cur_sheet_name,row,col+1).equals("") || PoiExcelTest.readExcelData(filelName,cur_sheet_name,row,col+1).equals("0")&& col < cur_col - 1) {
//////                                isStart = true;
//////                                tmptask.setUse_amount(String.valueOf(tmp_use + Double.parseDouble(PoiExcelTest.readExcelData(filelName,cur_sheet_name,row,col))));
//////                                tmptask.setDuration(tmptask.getDuration()+duration);
//////                                tmptask.setStart_date(tmp_start_date);
//////                                System.out.println(tmptask);
//////                                ganttTask.setDuration(0);
//////                            }
//////                            else{
//////                                tmp_use += Double.parseDouble(PoiExcelTest.readExcelData(filelName,cur_sheet_name,row,col));
//////                            }
//////                        }
//////                        else if(isStart == false && col == cur_col - 1){
//////                            tmptask.setStart_date(tmp_start_date);
//////                            tmptask.setUse_amount(String.valueOf(tmp_use));
//////                            System.out.println(tmptask);
//////                            isStart = true;
//////                        }
//////                        else {
//////                            ganttTask.setStart_date("");
//////                            ganttTask.setUse_amount("0");
//////                        }
////                        else if(!read_res.equals("") &&!read_res.equals("0")&& isStart == false){
////                            tmp_use += Double.parseDouble(read_res);
////                            tmptask2.setStart_date(PoiExcelTest.readExcelData(filelName,cur_sheet_name,0,col));
////                        }
////                        else if(col == cur_col -1){
////                            tmptask.setUse_amount(String.valueOf(tmp_use));
////                            String time_Str = tmptask2.getStart_date();
//////                            System.out.println(time_Str);
////                            time_date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.parseDouble(time_Str));
////                            int duration = (int) ((time_date.getTime() - (simpleDateFormat.parse((tmp_start_date)).getTime()))/86400000);
////                            tmptask.setDuration(tmptask.getDuration()+duration);
////                            tmptask.setStart_date(tmp_start_date);
////                            System.out.println(tmptask);
//////                            ganttTaskMapper.insertTask(tmptask);
////                        }
////                    }
////                }
////            }
////        }
//
//    }

    public static String readExcelData(String filelName, String sheetName, int rownum, int cellnum) throws Exception{
//        InputStream in = new FileInputStream(filelName);
//        workbook = new XSSFWorkbook(in);
        sheet = workbook.getSheet(sheetName);
        sheet.getRow(rownum).getCell(cellnum, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellType(CellType.STRING);
        String cellValue = sheet.getRow(rownum).getCell(cellnum).getStringCellValue();
       return  cellValue;
    }

    public static int readrowNum2(String filelName, String sheetName,int cur_mark) throws Exception{
        int count = 0;
//        InputStream in = new FileInputStream(filelName);
//        workbook = new XSSFWorkbook(in);
        sheet = workbook.getSheet(sheetName);
        int rowNum=sheet.getLastRowNum();
        for(int i = 1; i <= rowNum; i++){
            if(readExcelData(filelName,sheetName,i,cur_mark).equals("")){
                rowNum = i-1;
                break;
            }
        }
        return  rowNum;
    }

    public static int readrowNum(String filelName, String sheetName) throws Exception{
        int count = 0;
//        InputStream in = new FileInputStream(filelName);
//        workbook = new XSSFWorkbook(in);
        sheet = workbook.getSheet(sheetName);
        int rowNum=sheet.getLastRowNum();
        for(int i = 1; i < rowNum; i++){
//            System.out.println();
            if(readExcelData(filelName,sheetName,i,0).equals("")){
                count++;
            }
        }
        return  rowNum-count;
    }
    public static int readcolNum(String filelName, String sheetName) throws Exception{
//        InputStream in = new FileInputStream(filelName);
//        workbook = new XSSFWorkbook(in);
        sheet = workbook.getSheet(sheetName);
        int columnNum=sheet.getRow(0).getPhysicalNumberOfCells();
        return  columnNum;
    }

    public static void closeExcel() throws IOException {
        workbook.close();
    }
}