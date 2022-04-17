package com.tianma.yunying.service;

import com.tianma.yunying.entity.*;
//import com.tianma.yunying.mapper.GanttMapper;
import com.tianma.yunying.mapper.GanttTaskMapper;
import com.tianma.yunying.test.PoiExcelTest;
import com.tianma.yunying.util.Excel_Util;
import com.tianma.yunying.util.StringUtils;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import javax.swing.filechooser.FileSystemView;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.NumberFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

@Service
public class GanttTaskService {
    @Autowired
    GanttTaskMapper ganttTaskMapper;

    public List<GanttTask> getTask(int type){
        if(type == 3){
            return ganttTaskMapper.getTask_plan();
        }
        if(type == 6){
            return ganttTaskMapper.getTask_real();
        }
        else return ganttTaskMapper.getTask();
    }

    public List<RunGanttTask> getRunTask(){
        return ganttTaskMapper.getRunTask();
    }
    public List<MacroGantt> getMacroTask(){
        return ganttTaskMapper.getMacroTask();
    }
    public List<RunGanttTask> getRunTaskByPilot(String pilot){
        if(pilot.equals("all")){
            return ganttTaskMapper.getRunTask();
        }
        else {
            String a = pilot.split("pilot")[1];
            return ganttTaskMapper.getRunTaskByPilot(a);
        }
    }
    public List<GanttTask> getTaskByPilot(String pilot){
        if(pilot.equals("all")){
            return ganttTaskMapper.getTask();
        }
        else {
            return ganttTaskMapper.getTaskByPilot(pilot);
        }
    }
    public List<String> getAllDepart(){
        return ganttTaskMapper.getAllDepart();
    }
    public List<String> getAllProject(){
        return ganttTaskMapper.getAllProject();
    }
    public List<String> getAllPilot(String tablename){
        return ganttTaskMapper.getAllPilot(tablename);
    }
    public List<String> getAllTarget(){
        return ganttTaskMapper.getAllTarget();
    }
    public MacroGantt_Status getMacroStatus(){
        return ganttTaskMapper.getMacroStatus();
    }
    public List<GanttCapacity> getCapacity(){
        return ganttTaskMapper.getCapacity();
    }
    public void updateCapacity(GanttCapacity ganttCapacity){
        ganttTaskMapper.updateCapacity(ganttCapacity);
        try {
            updateByCapacity();
        } catch (ParseException e) {
            e.printStackTrace();
        }
    }
    public Result InsertTask() throws Exception {
//        String filelName = "D:\\CodePath\\test\\联排计划展示基础表 (1).xlsx";
//        HashMap<Integer,String> sheet_name = new HashMap<>();
//        sheet_name.put(1,"ARRAY厂计划");
//        sheet_name.put(2,"EVEN厂计划");
//        sheet_name.put(3,"TPOT厂计划 ");
//        sheet_name.put(4,"EAC厂计划");
//        sheet_name.put(5,"MODULE厂计划");
//        String cur_sheet_name = "";
//        int cur_row = 0;
//        int cur_col = 0;
//        String tmp_date = "";
//        GanttTask ganttTask = new GanttTask();
//        DateFormat format=new SimpleDateFormat("yyyy-MM-dd");
//        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
//        ganttTask.setOpen("true");
//        Boolean isStart = true;
//        for(int i = 1; i<=1; i++){
//            cur_sheet_name = sheet_name.get(i);
//            cur_row = PoiExcelTest.readrowNum(filelName,cur_sheet_name);
//            cur_col = PoiExcelTest.readcolNum(filelName,cur_sheet_name);
//            for(int row = 1; row <=cur_row; row++){
//                if(PoiExcelTest.readExcelData(filelName,cur_sheet_name,row,6).equals("计划产出"))
//                    ganttTask.setId(PoiExcelTest.readExcelData(filelName,cur_sheet_name,row,1)+"_OUT");
//                else
//                    ganttTask.setId(PoiExcelTest.readExcelData(filelName,cur_sheet_name,row,1));
//                ganttTask.setStart_date(null);
//                ganttTask.setDuration(0);
//                ganttTask.setUse_amount(null);
//                isStart = true;
//                for(int col = 0; col < cur_col; col++){
//                    if(col > 6){
//                        if(PoiExcelTest.readExcelData(filelName,cur_sheet_name,row,col) != "" && isStart)
//                        {
//                            String time_Str = PoiExcelTest.readExcelData(filelName,cur_sheet_name,0,col);
//                            Date time_date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.valueOf(time_Str));
//                            ganttTask.setStart_date(format.format(time_date));
//                            ganttTask.setUse_amount(PoiExcelTest.readExcelData(filelName,cur_sheet_name,row,col));
//                            tmp_date = format.format(time_date);
//                            ganttTask.setDuration(1);
//                            isStart = false;
//                        }
//                        else if(PoiExcelTest.readExcelData(filelName,cur_sheet_name,row,col) != "" && !isStart){
////                            Date date = simpleDateFormat.parse((ganttTask.getStart_date())).getTime();
//                            System.out.println(ganttTask.getStart_date());
//                            String time_Str = PoiExcelTest.readExcelData(filelName,cur_sheet_name,0,col);
//                            Date time_date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.valueOf(time_Str));
////                            ganttTask.setStart_date(format.format(time_date.toString()));
//                            System.out.println(format.format(time_date));
//                            int duration = (int) ((time_date.getTime() - (simpleDateFormat.parse((tmp_date)).getTime()))/86400000);
//                            tmp_date = format.format(time_date);
//                            System.out.println(duration);
//                            ganttTask.setDuration(duration);
//                            ganttTask.setUse_amount(PoiExcelTest.readExcelData(filelName,cur_sheet_name,row,col));
//                            System.out.println("----------------------------------");
//                        }
//                        else if(PoiExcelTest.readExcelData(filelName,cur_sheet_name,row,col) == "" && ganttTask.getStart_date() != null){
//                            ganttTask.setStart_date(null);
//                        }
//                        ganttTaskMapper.insertTask(ganttTask);
//                    }
//                }
//            }
//        }
        String filelName = "D:\\CodePath\\test\\联排计划展示基础表 (1).xlsx";
        String sheetname = "标签关系表";
        HashMap<String, List<String>> relation = new HashMap<>();
        List<String> list_string = new ArrayList<>();
        DateFormat format=new SimpleDateFormat("yyyy-MM-dd");
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
        Date time_date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.valueOf(PoiExcelTest.readExcelData(filelName,"ARRAY厂计划",0,7)));
        int r_row = PoiExcelTest.readrowNum(filelName,sheetname);
        GanttTask parent_task = new GanttTask();
        parent_task.setDuration(100);
        parent_task.setStart_date(format.format(time_date));
        parent_task.setColor("rgba(0,0,0,0.0)");
        parent_task.setOpen("true");
        parent_task.setParent("产品1");
        parent_task.setRender("split");
//        int col = readcolNum(filelName,sheetname);
        for(int i = 1; i <= r_row-1; i++){
            if(PoiExcelTest.readExcelData(filelName,sheetname,i,5).equals("是")){
                String key = PoiExcelTest.readExcelData(filelName,sheetname,i,7) +"\\*"+PoiExcelTest.readExcelData(filelName,sheetname,i,6);
                List<String> label = new LinkedList<>();
                for(int j = 0; j < 5; j++){
                    String tmp_string = PoiExcelTest.readExcelData(filelName,sheetname,i,j);
                    if(!list_string.contains(tmp_string)){
                        list_string.add(tmp_string);
                        label.add(PoiExcelTest.readExcelData(filelName,sheetname,i,j));
                    }
                }
                System.out.println(label);
                relation.put(key,label);
            }
        }
//        int tmp_var = 0;
//        for(Map.Entry<String, List<String>> entry : relation.entrySet()){
//            Iterator<String> it = entry.getValue().iterator();
//            while(it.hasNext()){//判断是否有迭代元素
////                System.out.println(it.next());//输出迭代出的元素
//                System.out.println("33333333333");
//                String a = it.next();
//                System.out.println("111111111"+a+"2222222222");
//                if(a.equals("")){
//                    break;
//                }
//                if(tmp_var == 0)
//                    parent_task.setGantt_type("Array计划");
//                else if(tmp_var == 1)
//                    parent_task.setGantt_type("EVEN计划");
//                else if(tmp_var == 2)
//                    parent_task.setGantt_type("TPOT计划");
//                else if(tmp_var == 3)
//                    parent_task.setGantt_type("EAC计划");
//                else if(tmp_var == 4)
//                    parent_task.setGantt_type("MODULE计划");
//                parent_task.setText(entry.getKey().split("\\*")[0].replace("\\",""));
//                parent_task.setDesc(entry.getKey().split("\\*")[1]);
//                parent_task.setId(a);
//                System.out.println(parent_task);
//                ganttTaskMapper.insertTask(parent_task);
//                tmp_var += 1;
//            }
//            tmp_var = 0;
//            System.out.println("---------------------");
//        }
        int tmp_var = 0;
        int tmp_count = 1;
        int tmp_factory_number = 1;
        Boolean is_test = false;
        Boolean is_Full = false;
        for(Map.Entry<String, List<String>> entry : relation.entrySet()){
            String tmp_number = "";
            System.out.println(is_test);

//            if(tmp_full_var == 0 && entry.getValue().size() == 5){
//                is_Full = false;
//                parent_task.setFactory_number(String.valueOf(tmp_factory_number));
//            }
//            else if(!is_Full && entry.getValue().size() <5){
//                parent_task.setFactory_number(String.valueOf(tmp_factory_number));
//            }
//            else if(entry.getValue().size() == 5){
//                tmp_factory_number += 1;
//                parent_task.setFactory_number(String.valueOf(tmp_factory_number));
//                is_Full = true;
//            }


//            if(entry.getValue().size() < 5 && !is_Full){
//                parent_task.setFactory_number(String.valueOf(tmp_factory_number));
//                is_test = true;
//            }
//            else if(entry.getValue().size() < 5 && is_Full){
//                is_Full = false;
//                tmp_factory_number += 1;
//                parent_task.setFactory_number(String.valueOf(tmp_factory_number));
//            }
//            else if(entry.getValue().size() == 5 && tmp_full_var == 1 && !is_Full){
//                tmp_factory_number += 1;
//                parent_task.setFactory_number(String.valueOf(tmp_factory_number));
////                tmp_factory_number += 1;
//            }
//            else if(entry.getValue().size() == 5 && tmp_full_var == 1 && is_Full){
//                tmp_factory_number += 1;
//                parent_task.setFactory_number(String.valueOf(tmp_factory_number));
//                is_Full = false;
////                tmp_factory_number += 1;
//            }
//            else if(entry.getValue().size() == 5 && tmp_full_var != 1 && !is_Full){
//                if(is_test){
//                    tmp_factory_number += 1;
//                }
//                parent_task.setFactory_number(String.valueOf(tmp_factory_number));
////                tmp_factory_number += 1;
//                tmp_full_var = 0;
//                is_Full = true;
//            }
//            else if(entry.getValue().size() == 5 && tmp_full_var != 1 && is_Full){
//                tmp_factory_number += 1;
//                parent_task.setFactory_number(String.valueOf(tmp_factory_number));
//                is_Full = false;
//            }
//            tmp_full_var += 1;

            if(entry.getValue().size() < 5 && !is_Full ){
                parent_task.setFactory_number(String.valueOf(tmp_factory_number));
            }
            else if(entry.getValue().size() == 5 && !is_Full&& !is_test){
                parent_task.setFactory_number(String.valueOf(tmp_factory_number));
                is_Full = true;
            }

            else if(entry.getValue().size() == 5 && is_Full){
                tmp_factory_number += 1;
                is_Full = false;
                is_test = true;

                parent_task.setFactory_number(String.valueOf(tmp_factory_number));
            }
            else if(entry.getValue().size() == 5 && is_test){
                is_test = false;
                tmp_factory_number += 1;
                parent_task.setFactory_number(String.valueOf(tmp_factory_number));
            }

//            tmp_factory_number += 1;
            if(entry.getValue().size() == 5){
                Iterator<String> it = entry.getValue().iterator();
                while(it.hasNext()) {//判断是否有迭代元素
                    String a = it.next();
                    if (a.equals("")) {
                        break;
                    }
                    if (tmp_var == 0) {
                        parent_task.setGantt_type("Array计划");
                        tmp_number += String.valueOf(tmp_count);
                        parent_task.setNumber(Integer.parseInt(tmp_number));
                    } else if (tmp_var == 1) {
                        parent_task.setGantt_type("EVEN计划");
                        tmp_number += String.valueOf(tmp_count);
                        parent_task.setNumber(Integer.parseInt(tmp_number));
                    } else if (tmp_var == 2) {
                        parent_task.setGantt_type("TPOT计划");
                        tmp_number += String.valueOf(tmp_count);
                        parent_task.setNumber(Integer.parseInt(tmp_number));
                    } else if (tmp_var == 3) {
                        parent_task.setGantt_type("EAC计划");
                        tmp_number += String.valueOf(tmp_count);
                        parent_task.setNumber(Integer.parseInt(tmp_number));
                    } else if (tmp_var == 4) {
                        parent_task.setGantt_type("MODULE计划");
                        tmp_number += String.valueOf(tmp_count);
                        parent_task.setNumber(Integer.parseInt(tmp_number));
                    }

                    parent_task.setText(entry.getKey().split("\\*")[0].replace("\\", ""));
                    parent_task.setDesc(entry.getKey().split("\\*")[1]);
                    parent_task.setId(a);
                    System.out.println(parent_task);
                    ganttTaskMapper.insertTask(parent_task);
                    tmp_var += 1;
                }
            }
            else {
                tmp_number = "";
                Iterator<String> it = entry.getValue().iterator();
                tmp_var = 5 - entry.getValue().size();
                System.out.println(tmp_var);
                Boolean flag = true;
                int count = 1;
                List<Integer> list_a = new ArrayList<>();
                while(it.hasNext()){
                    String a = it.next();
                    parent_task.setText(entry.getKey().split("\\*")[0].replace("\\", ""));
                    parent_task.setDesc(entry.getKey().split("\\*")[1]);
                    parent_task.setId(a);
                    if(flag) {
                        for (int i = 0; i <= tmp_var; i++) {
                            tmp_number += String.valueOf(tmp_count);
                        }
                        for (int i = 1; i <= entry.getValue().size(); i++) {
                            list_a.add(Integer.parseInt(tmp_number) + i);
                            System.out.println(Integer.parseInt(tmp_number) + i);
                        }
                        flag = false;
                    }
                    if(entry.getValue().size() - count == 0){
                        parent_task.setGantt_type("MODULE计划");
                    }
                    if(entry.getValue().size() - count == 1){
                        parent_task.setGantt_type("EAC计划");
                    }
                    if(entry.getValue().size() - count == 2){
                        parent_task.setGantt_type("TPOT计划");
                    }
                    if(entry.getValue().size() - count == 3){
                        parent_task.setGantt_type("EVEN计划");
                    }
                    parent_task.setNumber(list_a.get(count-1));
//                    System.out.println(parent_task);
                    ganttTaskMapper.insertTask(parent_task);
                    count++;
//                    }
                }

            }

            tmp_var = 0;
            tmp_number ="";
            tmp_count += 1;

            System.out.println("---------------------");
        }


        HashMap<Integer,String> sheet_name = new HashMap<>();
        sheet_name.put(1,"ARRAY厂计划");
        sheet_name.put(2,"EVEN厂计划");
        sheet_name.put(3,"TPOT厂计划 ");
        sheet_name.put(4,"EAC厂计划");
        sheet_name.put(5,"MODULE厂计划");
        String cur_sheet_name = "";
        int cur_row = 0;
        int cur_col = 0;
        GanttTask ganttTask = new GanttTask();
        GanttTask tmptask = new GanttTask();
        GanttTask tmptask2 = new GanttTask();
        String tmp_start_date = "";
        Double tmp_use = 0.0;
        ganttTask.setOpen("true");
        Boolean isStart = true;
        for(int i = 1; i<=5; i++){
            cur_sheet_name = sheet_name.get(i);
            cur_row = PoiExcelTest.readrowNum(filelName,cur_sheet_name);
            cur_col = PoiExcelTest.readcolNum(filelName,cur_sheet_name);
            for(int row = 1; row <=cur_row; row++){

                if(PoiExcelTest.readExcelData(filelName,cur_sheet_name,row,6).equals("计划投入")) {
                    ganttTask.setId(PoiExcelTest.readExcelData(filelName, cur_sheet_name, row, 1) + "计划投入");
                    if (i == 1)
                        ganttTask.setColor("rgba(255,165,0,0.4)");
                    else if (i == 2)
                        ganttTask.setColor("rgba(255,105,180,0.4)");
                    else if (i == 3)
                        ganttTask.setColor("rgba(0,255,0,0.4)");
                    else if (i == 4)
                        ganttTask.setColor("rgba(153,50,204,0.4)");
                    else if (i == 5)
                        ganttTask.setColor("rgba(210,180,140,0.4)");
                }
                else{
                    ganttTask.setId(PoiExcelTest.readExcelData(filelName,cur_sheet_name,row,1)+"计划产出");
                    ganttTask.setColor("rgba(0,128,128,0.4)");
                }
                ganttTask.setGantt_type(cur_sheet_name.split("厂")[0]+"计划");
                ganttTask.setText("test");
                ganttTask.setDuration(0);
                ganttTask.setProject_name("project_name");
                ganttTask.setParent(PoiExcelTest.readExcelData(filelName,cur_sheet_name,row,1));
                isStart = true;
                for(int col = 0; col < cur_col; col++){
                    if(col > 6){
                        String read_res =  PoiExcelTest.readExcelData(filelName,cur_sheet_name,row,col);
                        if(!read_res.equals("") &&!read_res.equals("0") && isStart && isStart == true)
                        {
                            String time_Str = PoiExcelTest.readExcelData(filelName,cur_sheet_name,0,col);
                            time_date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.valueOf(time_Str));
                            ganttTask.setStart_date(format.format(time_date));
                            ganttTask.setDuration(1);
                            ganttTask.setUse_amount(PoiExcelTest.readExcelData(filelName,cur_sheet_name,row,col));
                            tmp_start_date = ganttTask.getStart_date();
                            tmp_use = Double.parseDouble(ganttTask.getUse_amount());
                            tmptask = ganttTask;
                            tmptask2 = ganttTask;
                            tmptask2.setStart_date(time_Str);
                            isStart = false;
                        }
                        else if(!read_res.equals("") &&!read_res.equals("0")&& isStart == false){
                            tmp_use += Double.parseDouble(read_res);
                            tmptask2.setStart_date(PoiExcelTest.readExcelData(filelName,cur_sheet_name,0,col));
                        }
                        else if(col == cur_col -1){
                            tmptask.setUse_amount(String.valueOf(tmp_use));
                            String time_Str = tmptask2.getStart_date();
//                            System.out.println(time_Str);
                            time_date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.parseDouble(time_Str));
                            int duration = (int) ((time_date.getTime() - (simpleDateFormat.parse((tmp_start_date)).getTime()))/86400000);
                            tmptask.setDuration(tmptask.getDuration()+duration);
                            tmptask.setStart_date(tmp_start_date);
                            System.out.println(tmptask);
                            ganttTaskMapper.insertTask(tmptask);
                        }
                    }
                }
            }
        }
            return  new Result("200","success");
    }
    public Result InsertTask_real() throws Exception {
        String filelName = "D:\\CodePath\\test\\联排实际展示基础表.xlsx";
        String sheetname = "标签关系表";
        HashMap<String, List<String>> relation = new HashMap<>();
        DateFormat format=new SimpleDateFormat("yyyy-MM-dd");
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
        Date time_date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.valueOf(PoiExcelTest.readExcelData(filelName,"ARRAY厂实际",0,7)));
        int r_row = PoiExcelTest.readrowNum(filelName,sheetname);
        GanttTask parent_task = new GanttTask();
        parent_task.setDuration(100);
        parent_task.setStart_date(format.format(time_date));
        parent_task.setColor("rgba(0,0,0,0.0)");
        parent_task.setOpen("true");
        parent_task.setParent("产品1");
        parent_task.setRender("split");
//        int col = readcolNum(filelName,sheetname);
        for(int i = 1; i <= r_row; i++){
            if(PoiExcelTest.readExcelData(filelName,sheetname,i,5).equals("是")){
                String key = PoiExcelTest.readExcelData(filelName,sheetname,i,7) +"\\*"+PoiExcelTest.readExcelData(filelName,sheetname,i,6);
                List<String> label = new LinkedList<>();
                for(int j = 0; j < 5; j++){
                    label.add(PoiExcelTest.readExcelData(filelName,sheetname,i,j));
                }
                relation.put(key,label);
            }
        }
        int tmp_var = 0;
        int tmp_count = 1;
        for(Map.Entry<String, List<String>> entry : relation.entrySet()){

            String tmp_number = "";
            Iterator<String> it = entry.getValue().iterator();
            while(it.hasNext()){//判断是否有迭代元素
//                System.out.println(it.next());//输出迭代出的元素
                System.out.println("33333333333");
                String a = it.next();
                System.out.println("111111111"+a+"2222222222");
                if(a.equals("")){
                    break;
                }
                if(tmp_var == 0){
                    parent_task.setGantt_type("Array实际");
                    tmp_number += String.valueOf(tmp_count);
                    parent_task.setNumber(Integer.parseInt(tmp_number));
                }
                else if(tmp_var == 1){
                    parent_task.setGantt_type("EVEN实际");
                    tmp_number += String.valueOf(tmp_count);
                    parent_task.setNumber(Integer.parseInt(tmp_number));
                }
                else if(tmp_var == 2) {
                    parent_task.setGantt_type("TPOT实际");
                    tmp_number += String.valueOf(tmp_count);
                    parent_task.setNumber(Integer.parseInt(tmp_number));
                }
                else if(tmp_var == 3)
                {
                    parent_task.setGantt_type("EAC实际");
                    tmp_number += String.valueOf(tmp_count);
                    parent_task.setNumber(Integer.parseInt(tmp_number));
                }
                else if(tmp_var == 4)
                {
                    parent_task.setGantt_type("MODULE实际");
                    tmp_number += String.valueOf(tmp_count);
                    parent_task.setNumber(Integer.parseInt(tmp_number));
                }

                parent_task.setText(entry.getKey().split("\\*")[0].replace("\\",""));
                parent_task.setDesc(entry.getKey().split("\\*")[1]);
                parent_task.setId(a+"实际");
                System.out.println(parent_task);
                ganttTaskMapper.insertTask(parent_task);
                tmp_var += 1;

            }
            tmp_var = 0;
            tmp_number ="";
            tmp_count += 1;

            System.out.println("---------------------");
        }
        HashMap<Integer,String> sheet_name = new HashMap<>();
        sheet_name.put(1,"ARRAY厂实际");
        sheet_name.put(2,"EVEN厂实际");
        sheet_name.put(3,"TPOT厂实际 ");
        sheet_name.put(4,"EAC厂实际");
        sheet_name.put(5,"MODULE厂实际");
        String cur_sheet_name = "";
        int cur_row = 0;
        int cur_col = 0;
        GanttTask ganttTask = new GanttTask();
        GanttTask tmptask = new GanttTask();
        GanttTask tmptask2 = new GanttTask();
        String tmp_start_date = "";
        Double tmp_use = 0.0;
        ganttTask.setOpen("true");
        Boolean isStart = true;
        for(int i = 1; i<=5; i++){
            cur_sheet_name = sheet_name.get(i);
            cur_row = PoiExcelTest.readrowNum(filelName,cur_sheet_name);
            cur_col = PoiExcelTest.readcolNum(filelName,cur_sheet_name);
            for(int row = 1; row <=cur_row; row++){

                if(PoiExcelTest.readExcelData(filelName,cur_sheet_name,row,6).equals("实际投入")) {
                    ganttTask.setId(PoiExcelTest.readExcelData(filelName, cur_sheet_name, row, 1) + "实际投入");
                    if (i == 1)
                        ganttTask.setColor("rgba(255,165,0,0.4)");
                    else if (i == 2)
                        ganttTask.setColor("rgba(255,105,180,0.4)");
                    else if (i == 3)
                        ganttTask.setColor("rgba(0,255,0,0.4)");
                    else if (i == 4)
                        ganttTask.setColor("rgba(153,50,204,0.4)");
                    else if (i == 5)
                        ganttTask.setColor("rgba(210,180,140,0.4)");
                }
                else{
                    ganttTask.setId(PoiExcelTest.readExcelData(filelName,cur_sheet_name,row,1)+"实际产出");
                    ganttTask.setColor("rgba(0,128,128,0.4)");
                }
                ganttTask.setGantt_type(cur_sheet_name.split("厂")[0]+"实际");
                ganttTask.setText("test");
                ganttTask.setDuration(0);
                ganttTask.setProject_name("project_name");
                ganttTask.setParent(PoiExcelTest.readExcelData(filelName,cur_sheet_name,row,1)+"实际");
                isStart = true;
                for(int col = 0; col < cur_col; col++){
                    if(col > 6){
                        String read_res =  PoiExcelTest.readExcelData(filelName,cur_sheet_name,row,col);
                        if(!read_res.equals("") &&!read_res.equals("0") && isStart && isStart == true)
                        {
                            String time_Str = PoiExcelTest.readExcelData(filelName,cur_sheet_name,0,col);
                            time_date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.valueOf(time_Str));
                            ganttTask.setStart_date(format.format(time_date));
                            ganttTask.setDuration(1);
                            ganttTask.setUse_amount(PoiExcelTest.readExcelData(filelName,cur_sheet_name,row,col));
                            tmp_start_date = ganttTask.getStart_date();
                            tmp_use = Double.parseDouble(ganttTask.getUse_amount());
                            tmptask = ganttTask;
                            tmptask2 = ganttTask;
                            tmptask2.setStart_date(time_Str);
                            isStart = false;
                        }
//                        else if(!read_res.equals("") &&!read_res.equals("0") && isStart == false){
////                            System.out.println(read_res);
//                            String time_Str = PoiExcelTest.readExcelData(filelName,cur_sheet_name,0,col);
//                            Date time_date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.valueOf(time_Str));
//                            int duration = (int) ((time_date.getTime() - (simpleDateFormat.parse((tmp_start_date)).getTime()))/86400000);
//                            if(PoiExcelTest.readExcelData(filelName,cur_sheet_name,row,col+1).equals("") || PoiExcelTest.readExcelData(filelName,cur_sheet_name,row,col+1).equals("0")&& col < cur_col - 1) {
//                                isStart = true;
//                                tmptask.setUse_amount(String.valueOf(tmp_use + Double.parseDouble(PoiExcelTest.readExcelData(filelName,cur_sheet_name,row,col))));
//                                tmptask.setDuration(tmptask.getDuration()+duration);
//                                tmptask.setStart_date(tmp_start_date);
//                                System.out.println(tmptask);
//                                ganttTask.setDuration(0);
//                            }
//                            else{
//                                tmp_use += Double.parseDouble(PoiExcelTest.readExcelData(filelName,cur_sheet_name,row,col));
//                            }
//                        }
//                        else if(isStart == false && col == cur_col - 1){
//                            tmptask.setStart_date(tmp_start_date);
//                            tmptask.setUse_amount(String.valueOf(tmp_use));
//                            System.out.println(tmptask);
//                            isStart = true;
//                        }
//                        else {
//                            ganttTask.setStart_date("");
//                            ganttTask.setUse_amount("0");
//                        }
                        else if(!read_res.equals("") &&!read_res.equals("0")&& isStart == false){
                            tmp_use += Double.parseDouble(read_res);
                            tmptask2.setStart_date(PoiExcelTest.readExcelData(filelName,cur_sheet_name,0,col));
                        }
                        else if(col == cur_col -1){
                            tmptask.setUse_amount(String.valueOf(tmp_use));
                            String time_Str = tmptask2.getStart_date();
//                            System.out.println(time_Str);
                            time_date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.parseDouble(time_Str));
                            int duration = (int) ((time_date.getTime() - (simpleDateFormat.parse((tmp_start_date)).getTime()))/86400000);
                            tmptask.setDuration(tmptask.getDuration()+duration);
                            tmptask.setStart_date(tmp_start_date);
                            System.out.println(tmptask);
                            ganttTaskMapper.insertTask(tmptask);
                        }
                    }
                }
            }
        }
        return  new Result("200","success");
    }
    public Result new_task(String fileName) throws Exception {
        long start = System.currentTimeMillis();
//        String filelName = "D:\\CodePath\\test\\联排编码逻辑.xlsx";
//        String filelName = "D:\\CodePath\\test\\1_test.xlsx";
        ganttTaskMapper.deleteTask();
        PoiExcelTest.workbook = new XSSFWorkbook(fileName);
        HashMap<Integer, String> sheet_name = new HashMap<>();
        HashMap<String,Integer> sheet_factory = new HashMap<>();
        HashMap<String,String> sheet_relation = new HashMap<>();
        int rela_col = PoiExcelTest.readcolNum(fileName,"对应顺序表");
        int rela_row = PoiExcelTest.readrowNum(fileName,"对应顺序表");
        for(int i = 1; i <= rela_row; i++){
            for(int j = 0; j < rela_col; j++){
//                System.out.println(PoiExcelTest.readExcelData(filelName,"对应顺序表",i,j));
                String tmp = PoiExcelTest.readExcelData(fileName,"对应顺序表",i,2)+"^"+PoiExcelTest.readExcelData(fileName,"对应顺序表",i,0);
                sheet_relation.put(PoiExcelTest.readExcelData(fileName,"对应顺序表",i,1),tmp);
            }
        }
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
        int cur_in = 0;
        int cur_out = 0;
        int temp_i = 0;
        int temp_var = 0;
        DateFormat format=new SimpleDateFormat("yyyy/MM/dd");
        GanttTask ganttTask = new GanttTask();
        GanttTask parentTask = new GanttTask();
        GanttTask fatherTask = new GanttTask();
        GanttTask tmptask = new GanttTask();
        GanttTask tmptask2 = new GanttTask();
        String tmp_start_date = "";
        Double tmp_use = 0.0;
//        ganttTask.setOpen("true");
        Boolean isStart = true;
        for (int i = 1; i <= 5; i++) {
            cur_sheet_name = sheet_name.get(i);
            System.out.println(cur_sheet_name);
            cur_col = PoiExcelTest.readcolNum(fileName, cur_sheet_name);
            System.out.println("cur_col:" + cur_col);
            for (int j = 0; j <= cur_col; j++) {
                String tmp = PoiExcelTest.readExcelData(fileName, cur_sheet_name, 0, j);
                if (tmp.equals("Mark")) {
                    cur_mark = j;
                    cur_row = PoiExcelTest.readrowNum2(fileName, cur_sheet_name, cur_mark);
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

            temp_i++;
            Calendar cal = Calendar.getInstance();
            temp_var = 1;
            System.out.println("!!!!!!!!!!!!!!!!!!");
            System.out.println(sheet_factory);
            for (int row = 1; row <= cur_row; row++) {
                String var_temp = PoiExcelTest.readExcelData(fileName, cur_sheet_name, row, cur_mark);
                parentTask.setId(var_temp);
                parentTask.setColor("rgba(0,0,0,0)");
                parentTask.setDuration(100);
//                if(PoiExcelTest.readExcelData(fileName, cur_sheet_name, row, cur_target).contains("项目")){
//                    parentTask.setParent("项目送样和项目验证");
//                    parentTask.setText("项目送样和项目验证");
//                    if(!sheet_factory.containsKey("项目送样和项目验证")){
//                        fatherTask.setId("项目送样和项目验证");
//                        fatherTask.setText("项目送样和项目验证");
//                        fatherTask.setColor("rgba(0,0,0,0)");
//                        fatherTask.setOpen("true");
//                        fatherTask.setStart_date(String.valueOf(format.format(org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.valueOf(PoiExcelTest.readExcelData(fileName, cur_sheet_name, 0, cur_type+1))))));
//                        ganttTaskMapper.insertTask(fatherTask);
//                        sheet_factory.put("项目送样和项目验证",temp_var);
//                        temp_var++;
//                    }
//                }
//                else {
//                    String[] a = PoiExcelTest.readExcelData(fileName, cur_sheet_name, row, cur_target).split(" ");
//                    parentTask.setParent(a[a.length-1]);
//                    parentTask.setText(a[a.length-1]);
//                    if(!sheet_factory.containsKey(a[a.length-1])&&!a[a.length-1].equals("")){
//                        fatherTask.setId(a[a.length-1]);
//                        fatherTask.setText(a[a.length-1]);
//                        fatherTask.setColor("rgba(0,0,0,0)");
//                        fatherTask.setOpen("true");
//                        fatherTask.setStart_date(String.valueOf(format.format(org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.valueOf(PoiExcelTest.readExcelData(fileName, cur_sheet_name, 0, cur_type+1))))));
//                        ganttTaskMapper.insertTask(fatherTask);
//                        sheet_factory.put(a[a.length-1],temp_var);
//                        temp_var++;
//                    }
//                }
                String sub_string = PoiExcelTest.readExcelData(fileName, cur_sheet_name, row, cur_mark).substring(0,3);
                parentTask.setParent(sheet_relation.get(sub_string).split("\\^")[0]);
                parentTask.setText(PoiExcelTest.readExcelData(fileName, cur_sheet_name, row, cur_target));
                fatherTask.setId(sheet_relation.get(sub_string).split("\\^")[0]);
                fatherTask.setText(sheet_relation.get(sub_string).split("\\^")[0]);
                fatherTask.setColor("rgba(0,0,0,0)");
                fatherTask.setNumber(Integer.parseInt(sheet_relation.get(sub_string).split("\\^")[1]));
                fatherTask.setOpen("true");
                fatherTask.setStart_date(String.valueOf(format.format(org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.valueOf(PoiExcelTest.readExcelData(fileName, cur_sheet_name, 0, cur_type+1))))));
                ganttTaskMapper.insertTask(fatherTask);

                parentTask.setRender("split");
//                parentTask.setOpen("true");

                parentTask.setStart_date(String.valueOf(format.format(org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.valueOf(PoiExcelTest.readExcelData(fileName, cur_sheet_name, 0, cur_type+1))))));
//                System.out.println(format.parse(parentTask.getStart_date()));
                cal.setTime(format.parse(parentTask.getStart_date()));
                cal.add(Calendar.DATE,parentTask.getDuration()-1);
                parentTask.setEnd_date_text(String.valueOf(format.format(cal.getTime())));
                ganttTask.setParent(parentTask.getId());
                if (PoiExcelTest.readExcelData(fileName, cur_sheet_name, row, cur_type).equals("IN")) {
                    ganttTask.setId(var_temp + "投入");
                    if (i == 1) {
                        ganttTask.setColor("rgba(255,165,0,0.5)");
                        parentTask.setGantt_type("ARRAY计划");
                        System.out.println(parentTask);
                    } else if (i == 2) {
//                        ganttTask.setColor("rgba(255,105,180,0.4)");
                        ganttTask.setColor("rgba(255,165,0,0.5)");
                        parentTask.setGantt_type("EVEN计划");
                        System.out.println(parentTask);
                    } else if (i == 3) {
//                        ganttTask.setColor("rgba(0,255,0,0.4)");
                        ganttTask.setColor("rgba(255,165,0,0.5)");
                        parentTask.setGantt_type("TPOT计划");
                        System.out.println(parentTask);
                    } else if (i == 4) {
//                        ganttTask.setColor("rgba(153,50,204,0.4)");
                        ganttTask.setColor("rgba(255,165,0,0.5)");
                        parentTask.setGantt_type("EAC计划");
                        System.out.println(parentTask);
                    } else if (i == 5) {
//                        ganttTask.setColor("rgba(210,180,140,0.4)");
                        ganttTask.setColor("rgba(255,165,0,0.5)");
                        parentTask.setGantt_type("MODULE计划");
                        System.out.println(parentTask);
                    }

                } else {
                    ganttTask.setId(PoiExcelTest.readExcelData(fileName, cur_sheet_name, row, cur_mark) + "产出");
                    ganttTask.setColor("rgba(192,192,192,0.5)");
                }
                if(PoiExcelTest.readExcelData(fileName, cur_sheet_name, row, cur_type).equals("IN")){
//                    parentTask.setFactory_number(String.valueOf(temp_i));
                    parentTask.setFactory_number(String.valueOf(sheet_factory.get(parentTask.getText())));
                    String tmp_str = "";
                    for(int tmp =0; tmp < temp_i; tmp++)
                        tmp_str += String.valueOf(temp_i);
                    parentTask.setNumber(Integer.parseInt(tmp_str));
                    ganttTaskMapper.insertTask(parentTask);
                }

                ganttTask.setDuration(0);
                ganttTask.setStart_date("");
                ganttTask.setUse_amount("");
                isStart = true;
                for (int col = cur_type+1; col <= cur_col; col++) {
                    String read_res = PoiExcelTest.readExcelData(fileName, cur_sheet_name, row, col);
                    if (!read_res.equals("") && !read_res.equals("0") && isStart && isStart == true) {
                        System.out.println("11111111111111111  "+read_res);
                        String time_Str = PoiExcelTest.readExcelData(fileName, cur_sheet_name, 0, col);
                        Date time_date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.valueOf(time_Str));
                        ganttTask.setStart_date(format.format(time_date));
                        ganttTask.setDuration(1);
                        ganttTask.setUse_amount(PoiExcelTest.readExcelData(fileName, cur_sheet_name, row, col));
                        tmp_start_date = ganttTask.getStart_date();
                        tmp_use = Double.parseDouble(ganttTask.getUse_amount());
                        tmptask = ganttTask;
                        tmptask2 = ganttTask;
                        tmptask2.setStart_date(time_Str);
                        isStart = false;
                    } else if (!read_res.equals("") && !read_res.equals("0") && isStart == false) {
                        tmp_use += Double.parseDouble(read_res);
                        tmptask2.setStart_date(PoiExcelTest.readExcelData(fileName, cur_sheet_name, 0, col));
                    } else if (col == cur_col - 1) {
                        tmptask.setUse_amount(String.valueOf(tmp_use));
                        String time_Str = tmptask2.getStart_date();
                        Date time_date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.parseDouble(time_Str));
                        int duration = (int) ((time_date.getTime() - (format.parse((tmp_start_date)).getTime())) / 86400000);
                        tmptask.setDuration(tmptask.getDuration() + duration);
                        tmptask.setStart_date(tmp_start_date);
                        System.out.println(tmptask);
                        cal.setTime(format.parse(tmptask.getStart_date()));
                        cal.add(Calendar.DATE,tmptask.getDuration()-1);
                        tmptask.setEnd_date_text(String.valueOf(format.format(cal.getTime())));
                        if(!parentTask.getText().equals("")){
                            tmptask.setText(parentTask.getText());
                        }
                        else {
                            tmptask.setText(fatherTask.getText());
                        }
//                        tmptask.setOpen("true");
                        ganttTaskMapper.insertTask(tmptask);
                    }
                }
//                System.out.println(ganttTask);
            }
        }
        long end = System.currentTimeMillis();
        System.out.println("程序运行时间："+(end-start)+"ms");
        PoiExcelTest.closeExcel();
//        File file = new File(fileName);
//        file.delete();
        return new Result("200","success");
    }
    public Result new_task_real(String fileName) throws Exception{
        long start = System.currentTimeMillis();

//        String filelName = "D:\\CodePath\\test\\联排编码逻辑.xlsx";
//        String filelName = "D:\\CodePath\\test\\2_test.xlsx";
        PoiExcelTest.workbook = new XSSFWorkbook(fileName);
        HashMap<Integer, String> sheet_name = new HashMap<>();
        HashMap<String,Integer> sheet_factory = new HashMap<>();
        sheet_name.put(1, "ARRAY实际");
        sheet_name.put(2, "EVEN实际");
        sheet_name.put(3, "TPOT实际");
        sheet_name.put(4, "EAC实际");
        sheet_name.put(5, "MODULE实际");
        HashMap<String,String> sheet_relation = new HashMap<>();
        int rela_col = PoiExcelTest.readcolNum(fileName,"对应顺序表");
        int rela_row = PoiExcelTest.readrowNum(fileName,"对应顺序表");
        for(int i = 1; i <= rela_row; i++){
            for(int j = 0; j < rela_col; j++){
//                System.out.println(PoiExcelTest.readExcelData(filelName,"对应顺序表",i,j));
                String tmp = PoiExcelTest.readExcelData(fileName,"对应顺序表",i,2)+"^"+PoiExcelTest.readExcelData(fileName,"对应顺序表",i,0);
                sheet_relation.put(PoiExcelTest.readExcelData(fileName,"对应顺序表",i,1),tmp);
            }
        }

        String cur_sheet_name = "";
        int cur_row = 0;
        int cur_col = 0;
        int cur_type=0;
        int cur_mark=0;
        int cur_target=0;
        int cur_in = 0;
        int cur_out = 0;
        int temp_i = 0;
        int temp_var = 0;
        DateFormat format=new SimpleDateFormat("yyyy/MM/dd");
        GanttTask ganttTask = new GanttTask();
        GanttTask parentTask = new GanttTask();
        GanttTask fatherTask = new GanttTask();
        GanttTask tmptask = new GanttTask();
        GanttTask tmptask2 = new GanttTask();
        String tmp_start_date = "";
        Double tmp_use = 0.0;
//        ganttTask.setOpen("true");
        Boolean isStart = true;
        for (int i = 1; i <= 5; i++) {
            cur_sheet_name = sheet_name.get(i);
            System.out.println(cur_sheet_name);
            cur_col = PoiExcelTest.readcolNum(fileName, cur_sheet_name);
            System.out.println("cur_col:" + cur_col);
            for (int j = 0; j <= cur_col; j++) {
                String tmp = PoiExcelTest.readExcelData(fileName, cur_sheet_name, 0, j);
                if (tmp.equals("Mark")) {
                    cur_mark = j;
                    cur_row = PoiExcelTest.readrowNum2(fileName, cur_sheet_name, cur_mark);
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

            temp_i++;
            Calendar cal = Calendar.getInstance();
            temp_var = 1;
            System.out.println("!!!!!!!!!!!!!!!!!!");
            System.out.println(sheet_factory);
            for (int row = 1; row <= cur_row; row++) {
                String var_temp = PoiExcelTest.readExcelData(fileName, cur_sheet_name, row, cur_mark);
                parentTask.setId(var_temp+"实际");
                parentTask.setColor("rgba(0,0,0,0)");
                parentTask.setDuration(100);

//                if(PoiExcelTest.readExcelData(fileName, cur_sheet_name, row, cur_target).contains("项目")){
//                    parentTask.setParent("项目送样和项目验证");
//                    parentTask.setText("项目送样和项目验证");
//                    if(!sheet_factory.containsKey("项目送样和项目验证")){
//                        sheet_factory.put("项目送样和项目验证",temp_var);
//                        temp_var++;
//                    }
//                }
//                else {
//                    String[] a = PoiExcelTest.readExcelData(fileName, cur_sheet_name, row, cur_target).split(" ");
//                    parentTask.setParent(a[a.length-1]);
//                    parentTask.setText(a[a.length-1]);
//                    if(!sheet_factory.containsKey(a[a.length-1])&&!a[a.length-1].equals("")){
//                        sheet_factory.put(a[a.length-1],temp_var);
//                        temp_var++;
//                    }
//                }
                String sub_string = PoiExcelTest.readExcelData(fileName, cur_sheet_name, row, cur_mark).substring(0,3);
                parentTask.setParent(sheet_relation.get(sub_string).split("\\^")[0]);
                parentTask.setText(PoiExcelTest.readExcelData(fileName, cur_sheet_name, row, cur_target));
                parentTask.setRender("split");
//                parentTask.setOpen("true");

                parentTask.setStart_date(String.valueOf(format.format(org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.valueOf(PoiExcelTest.readExcelData(fileName, cur_sheet_name, 0, cur_type+1))))));
//                System.out.println(format.parse(parentTask.getStart_date()));
                cal.setTime(format.parse(parentTask.getStart_date()));
                cal.add(Calendar.DATE,parentTask.getDuration()-1);
                parentTask.setEnd_date_text(String.valueOf(format.format(cal.getTime())));
                ganttTask.setParent(parentTask.getId());
                if (PoiExcelTest.readExcelData(fileName, cur_sheet_name, row, cur_type).equals("IN")) {
                    ganttTask.setId(var_temp + "实际投入");
                    if (i == 1) {
                        ganttTask.setColor("rgba(0,191,255,0.8)");
                        parentTask.setGantt_type("ARRAY实际");
                        System.out.println(parentTask);
                    } else if (i == 2) {
                        ganttTask.setColor("rgba(0,191,255,0.8)");
                        parentTask.setGantt_type("EVEN实际");
                        System.out.println(parentTask);
                    } else if (i == 3) {
                        ganttTask.setColor("rgba(0,191,255,0.8)");
                        parentTask.setGantt_type("TPOT实际");
                        System.out.println(parentTask);
                    } else if (i == 4) {
                        ganttTask.setColor("rgba(0,191,255,0.8)");
                        parentTask.setGantt_type("EAC实际");
                        System.out.println(parentTask);
                    } else if (i == 5) {
                        ganttTask.setColor("rgba(0,191,255,0.8)");
                        parentTask.setGantt_type("MODULE实际");
                        System.out.println(parentTask);
                    }

                } else {
                    ganttTask.setId(PoiExcelTest.readExcelData(fileName, cur_sheet_name, row, cur_mark) + "实际产出");
//                    ganttTask.setColor("rgba(0,128,128,0.4)");
                    ganttTask.setColor("rgba(192,192,192,0.5)");
                }
                if(PoiExcelTest.readExcelData(fileName, cur_sheet_name, row, cur_type).equals("IN")){
//                    parentTask.setFactory_number(String.valueOf(temp_i));
                    parentTask.setFactory_number(String.valueOf(sheet_factory.get(parentTask.getText())));
                    String tmp_str = "";
                    for(int tmp =0; tmp < temp_i; tmp++)
                        tmp_str += String.valueOf(temp_i);
                    parentTask.setNumber(Integer.parseInt(tmp_str)+1);
                    if(!parentTask.getStart_date().equals(""))
                        ganttTaskMapper.insertTask(parentTask);
                }
                ganttTask.setDuration(0);
                ganttTask.setStart_date("");
                ganttTask.setUse_amount("");
                isStart = true;
                for (int col = cur_type+1; col <= cur_col; col++) {
//                    System.out.println(col);
                    String read_res = PoiExcelTest.readExcelData(fileName, cur_sheet_name, row, col);
                    if (!read_res.equals("") && !read_res.equals("0") && isStart && isStart == true) {
                        System.out.println("11111111111111111  "+read_res);
                        String time_Str = PoiExcelTest.readExcelData(fileName, cur_sheet_name, 0, col);
                        Date time_date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.valueOf(time_Str));
                        ganttTask.setStart_date(format.format(time_date));
                        ganttTask.setDuration(1);
                        ganttTask.setUse_amount(PoiExcelTest.readExcelData(fileName, cur_sheet_name, row, col));
                        tmp_start_date = ganttTask.getStart_date();
                        tmp_use = Double.parseDouble(ganttTask.getUse_amount());
                        tmptask = ganttTask;
                        tmptask2 = ganttTask;
                        tmptask2.setStart_date(time_Str);
                        isStart = false;
                    } else if (!read_res.equals("") && !read_res.equals("0") && isStart == false) {
                        tmp_use += Double.parseDouble(read_res);
                        tmptask2.setStart_date(PoiExcelTest.readExcelData(fileName, cur_sheet_name, 0, col));
                    } else if (col == cur_col - 1) {
                        tmptask.setUse_amount(String.valueOf(tmp_use));
                        String time_Str = tmptask2.getStart_date();
//                            System.out.println(time_Str);
                        Date time_date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.parseDouble(time_Str));
                        int duration = (int) ((time_date.getTime() - (format.parse((tmp_start_date)).getTime())) / 86400000);
                        tmptask.setDuration(tmptask.getDuration() + duration);
                        tmptask.setStart_date(tmp_start_date);
                        System.out.println(tmptask);
                        cal.setTime(format.parse(tmptask.getStart_date()));
                        cal.add(Calendar.DATE,tmptask.getDuration()-1);
                        tmptask.setEnd_date_text(String.valueOf(format.format(cal.getTime())));
                        if(!parentTask.getText().equals("")){
                            tmptask.setText(parentTask.getText());
                        }
                        else {
                            String temp_text = PoiExcelTest.readExcelData(fileName, cur_sheet_name, row, cur_target);
                            if(!temp_text.equals(""))
                                tmptask.setText(temp_text);
                            else
                                tmptask.setText(PoiExcelTest.readExcelData(fileName, cur_sheet_name, row-1, cur_target));
                        }
//                        tmptask.setOpen("true");

                        ganttTaskMapper.insertTask(tmptask);
                    }
                }
//                System.out.println(ganttTask);
            }
        }
        long end = System.currentTimeMillis();
        System.out.println("程序运行时间："+(end-start)+"ms");
        PoiExcelTest.closeExcel();
//        File file = new File(fileName);
//        file.delete();
        return new Result("200","success");
    }
    public Result InsertTask_new(String fileName)throws Exception {
        String sheetName = "编码规则";
        Excel_Util.workbook = new XSSFWorkbook(fileName);
        ganttTaskMapper.deleteTask();
        ganttTaskMapper.deleteMacroTask();
        int rela_row = Excel_Util.workbook.getSheet(sheetName).getLastRowNum();
        HashMap<String, String> code_Fir = new HashMap<>();
        HashMap<String, String> code_Sec = new HashMap<>();
        HashMap<String, String> code_Thi = new HashMap<>();
        HashMap<String, String> code_Fou = new HashMap<>();
        DateFormat format = new SimpleDateFormat("yyyy/MM/dd");
        String cur_sheet_name = "";
        GanttTask ganttTask = new GanttTask();
        GanttTask parentTask = new GanttTask();
        GanttTask fatherTask = new GanttTask();
        GanttTask tmptask = new GanttTask();
        GanttTask tmptask2 = new GanttTask();

        HashMap<Integer,String>facotry_name = new HashMap<>();
        facotry_name.put(1,"Array");
        facotry_name.put(2,"EVEN");
        facotry_name.put(3,"TPOT");
        facotry_name.put(4,"EAC");
        facotry_name.put(5,"MODULE");
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

        HashMap<String,String> update_plan_end = new HashMap<>();

        String tmp_str = "*";
        int tmp_int = 1;
        //将编码规则读入字典
        while (!tmp_str.equals("") && tmp_int <= rela_row) {
            tmp_str = Excel_Util.readExcelData(fileName, sheetName, tmp_int, 0);
            System.out.println(tmp_str);
            System.out.println(Excel_Util.readExcelData(fileName, sheetName, tmp_int, 1));
            code_Fir.put(Excel_Util.readExcelData(fileName, sheetName, tmp_int, 1), Excel_Util.readExcelData(fileName, sheetName, tmp_int, 0));
            tmp_int++;
        }
        tmp_int = 1;
        tmp_str = "*";
        while (!tmp_str.equals("") && tmp_int <= rela_row) {
            tmp_str = Excel_Util.readExcelData(fileName, sheetName, tmp_int, 2);
            System.out.println(tmp_str);
            System.out.println(Excel_Util.readExcelData(fileName, sheetName, tmp_int, 3));
            code_Sec.put(Excel_Util.readExcelData(fileName, sheetName, tmp_int, 3), Excel_Util.readExcelData(fileName, sheetName, tmp_int, 2));
            tmp_int++;
        }
        System.out.println(code_Sec);
        tmp_int = 1;
        tmp_str = "*";
        while (!tmp_str.equals("") && tmp_int <= rela_row) {
            tmp_str = Excel_Util.readExcelData(fileName, sheetName, tmp_int, 4);
            System.out.println(tmp_str);
            System.out.println(Excel_Util.readExcelData(fileName, sheetName, tmp_int, 5));
            code_Thi.put(Excel_Util.readExcelData(fileName, sheetName, tmp_int, 5), Excel_Util.readExcelData(fileName, sheetName, tmp_int, 4));
            tmp_int++;
        }
        tmp_int = 1;
        tmp_str = "*";
        while (!tmp_str.equals("") && tmp_int <= rela_row) {
            tmp_str = Excel_Util.readExcelData(fileName, sheetName, tmp_int, 6);
            System.out.println(tmp_str);
            System.out.println(Excel_Util.readExcelData(fileName, sheetName, tmp_int, 7));
            code_Fou.put(Excel_Util.readExcelData(fileName, sheetName, tmp_int, 7), Excel_Util.readExcelData(fileName, sheetName, tmp_int, 6));
            tmp_int++;
        }
        System.out.println(code_Thi);

        for (int sheet_index = 1; sheet_index <= 5; sheet_index++) {
            //获得当前表的各种信息
            cur_sheet_name = sheet_name.get(sheet_index);
            int cur_row = Excel_Util.readrowNum(fileName, cur_sheet_name);
            int cur_col = Excel_Util.readcolNum(fileName, cur_sheet_name);
            int cur_mark = Excel_Util.readWantCol(fileName, cur_sheet_name, 0, "Mark");
            int cur_target = Excel_Util.readWantCol(fileName, cur_sheet_name, 0, "投入目的");
            int cur_type = Excel_Util.readWantCol(fileName, cur_sheet_name, 0, "IN/OUT");
            System.out.println("!!!!!!!1");
            System.out.println(cur_type);
            //遍历五个工厂
            for (int row = 1; row <= cur_row; row++) {
                parentTask.setColor("rgba(0,0,0,0)");
                parentTask.setDuration(100);
                //切mark前四位来判断为父task
                String sub_string = Excel_Util.readExcelData(fileName, cur_sheet_name, row, cur_mark).substring(0, 4);
//                    System.out.println(code_Sec.get(sub_string.substring(1,2)));
//                parentTask.setId(code_Fir.get(sub_string.substring(0, 1)) + code_Sec.get(sub_string.substring(1, 2)) + code_Thi.get(sub_string.substring(2, 3)) + code_Fou.get(sub_string.substring(3, 4)));
//                parentTask.setText(code_Sec.get(sub_string.substring(1, 2)) + code_Fou.get(sub_string.substring(3, 4)) + code_Thi.get(sub_string.substring(2, 3)));
                parentTask.setRender("split");
                fatherTask.setId(code_Fir.get(sub_string.substring(0, 1)) + code_Sec.get(sub_string.substring(1, 2)) + code_Thi.get(sub_string.substring(2, 3)) + code_Fou.get(sub_string.substring(3, 4)));
                fatherTask.setText(fatherTask.getId());
                fatherTask.setColor("rgba(0,0,0,0)");
//                    fatherTask.setNumber(Integer.parseInt(sheet_relation.get(sub_string).split("\\^")[1]));
                fatherTask.setOpen("true");

                fatherTask.setPilot(code_Thi.get(sub_string.substring(2, 3)));

                ganttTaskMapper.insertTask(fatherTask);
                parentTask.setStart_date(String.valueOf(format.format(org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.valueOf(Excel_Util.readExcelData(fileName, cur_sheet_name, 0, cur_type + 1))))));
                cal.setTime(format.parse(parentTask.getStart_date()));
                cal.add(Calendar.DATE, parentTask.getDuration() - 1);
                parentTask.setEnd_date_text(String.valueOf(format.format(cal.getTime())));
//                    System.out.println(parentTask);

                if (Excel_Util.readExcelData(fileName, cur_sheet_name, row, cur_type).equals("IN")) {
                    ganttTask.setId(Excel_Util.readExcelData(fileName, cur_sheet_name, row, cur_mark) + "投入");
//                    macroGantt.setColor("rgba(255,165,0,0.5)");
//                    macroGantt.setId(ganttTask.getId());
//                    macroGantt.setText(Excel_Util.readExcelData(fileName, cur_sheet_name, row, cur_mark));

                    if (sheet_index == 1) {
                        ganttTask.setColor("rgba(255,165,0,0.5)");
                        parentTask.setGantt_type("ARRAY计划");
                        parentTask.setNumber(1);
//                        macroGantt.setParent("Array投入");
                    } else if (sheet_index == 2) {
                        ganttTask.setColor("rgba(255,165,0,0.5)");
                        parentTask.setGantt_type("EVEN计划");
                        parentTask.setNumber(2);
//                        macroGantt.setParent("EVEN投入");
                    } else if (sheet_index == 3) {
                        ganttTask.setColor("rgba(255,165,0,0.5)");
                        parentTask.setGantt_type("TPOT计划");
                        parentTask.setNumber(3);
//                        macroGantt.setParent("TPOT投入");
                    } else if (sheet_index == 4) {
                        ganttTask.setColor("rgba(255,165,0,0.5)");
                        parentTask.setGantt_type("EAC计划");
                        parentTask.setNumber(4);
//                        macroGantt.setParent("EAC投入");
                    } else if (sheet_index == 5) {
                        ganttTask.setColor("rgba(255,165,0,0.5)");
                        parentTask.setGantt_type("MODULE计划");
                        parentTask.setNumber(5);
//                        macroGantt.setParent("MODULE投入");
                    }

                } else {
                    ganttTask.setId(Excel_Util.readExcelData(fileName, cur_sheet_name, row, cur_mark) + "产出");
//                    macroGantt.setId(ganttTask.getId());
                    ganttTask.setColor("rgba(192,192,192,0.5)");
//                    macroGantt.setColor(ganttTask.getColor());
//                    if(sheet_index == 1){
//                        macroGantt.setParent("Array产出");
//                    }
//                    if(sheet_index == 2){
//                        macroGantt.setParent("EVEN产出");
//                    }
//                    if(sheet_index == 3){
//                        macroGantt.setParent("TPOT产出");
//                    }
//                    if(sheet_index == 4){
//                        macroGantt.setParent("EAC产出");
//                    }
//                    if(sheet_index == 5){
//                        macroGantt.setParent("MODULE产出");
//                    }
                }

                //创建parent任务
                if (Excel_Util.readExcelData(fileName, cur_sheet_name, row, cur_type).equals("IN")) {
                    parentTask.setId(Excel_Util.readExcelData(fileName,cur_sheet_name,row,cur_mark));
                    parentTask.setPilot(code_Thi.get(String.valueOf(parentTask.getId().charAt(2))));
                    parentTask.setText(Excel_Util.readExcelData(fileName,cur_sheet_name,row,cur_target));
                    parentTask.setParent(fatherTask.getId());
                    System.out.println(parentTask);
                    ganttTaskMapper.insertTask(parentTask);
                }

                ganttTask.setParent(parentTask.getId());
                ganttTask.setDuration(0);
                ganttTask.setStart_date("");
                ganttTask.setUse_amount("");
                ganttTask.setPilot(parentTask.getPilot());
                isStart = true;
                int cur_index = 0;
                for (int col = cur_type + 1; col <= cur_col; col++) {
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
                        if (col == cur_col - 1) {
                            String time_Str = Excel_Util.readExcelData(fileName, cur_sheet_name, 0, cur_index);
                            Date time_date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.valueOf(time_Str));
                            tmptask.setEnd_date_text(format.format(time_date));
                            if(update_plan_end.containsKey(tmptask.getParent()+"产出")&& !update_plan_end.get(tmptask.getParent()+"产出").equals("0")){
                                if(format.parse(update_plan_end.get(tmptask.getParent()+"产出")).getTime() < format.parse(tmptask.getEnd_date_text()).getTime()){
                                    update_plan_end.put(tmptask.getParent()+"产出",tmptask.getEnd_date_text());
                                }
                            }
                            else{
                                update_plan_end.put(tmptask.getParent()+"产出",tmptask.getEnd_date_text());
                            }
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
                        if (StringUtils.isNumeric(str_time) && !str_time.equals("")) {
                            String time_Str_end = Excel_Util.readExcelData(fileName, cur_sheet_name, 0, cur_index);
                            Date time_date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.valueOf(time_Str_end));
                            tmptask.setEnd_date_text(format.format(time_date));
                            if(tmptask.getId().contains("产出")){
                                if(update_plan_end.containsKey(tmptask.getParent()+"产出") && !update_plan_end.get(tmptask.getParent()+"产出").equals("0")){
                                    if(format.parse(update_plan_end.get(tmptask.getParent()+"产出")).getTime() < format.parse(tmptask.getEnd_date_text()).getTime()){
                                        update_plan_end.put(tmptask.getParent()+"产出",tmptask.getEnd_date_text());
                                    }
                                }
                                else{
                                    update_plan_end.put(tmptask.getParent()+"产出",tmptask.getEnd_date_text());
                                }

                            }
                        } else
                        {   tmptask.setEnd_date_text("0");
                            if(update_plan_end.containsKey(tmptask.getParent()+"产出")&& !update_plan_end.get(tmptask.getParent()+"产出").equals("0")){
                                if(format.parse(update_plan_end.get(tmptask.getParent()+"产出")).getTime() < format.parse(tmptask.getEnd_date_text()).getTime()){
                                    update_plan_end.put(tmptask.getParent()+"产出",tmptask.getEnd_date_text());
                                }
                            }
                            else{
                                update_plan_end.put(tmptask.getParent()+"产出",tmptask.getEnd_date_text());
                            }
                        }

                        if (!parentTask.getText().equals("")) {
                            tmptask.setText(parentTask.getText());
                        } else {
                            tmptask.setText(fatherTask.getText());
                        }
                        tmptask.setCode1(code_Fir.get(String.valueOf(tmptask.getId().charAt(0))));
                        tmptask.setCode2(code_Sec.get(String.valueOf(tmptask.getId().charAt(1))));
                        tmptask.setCode3(code_Thi.get(String.valueOf(tmptask.getId().charAt(2))));
                        tmptask.setCode4(code_Fou.get(String.valueOf(tmptask.getId().charAt(3))));

                        ganttTaskMapper.insertTask(tmptask);
                        System.out.println(tmptask);
                    }
                }

            }
            MacroGantt_Status macroGantt_status = new MacroGantt_Status();
            macroGantt_status.setTarget("-1");
            macroGantt_status.setProject("-1");
            macroGantt_status.setPilot("-1");
            macroGantt_status.setDepart("-1");
            setMacroGantt(macroGantt_status);
            //修改计划投入的结束时间至产出的结束时间
            for (Map.Entry<String, String> entry : update_plan_end.entrySet()) {
                if(!entry.getValue().equals("0")){
                    String str = entry.getKey().replace("产出","投入");
                    System.out.println(str);
                    ganttTaskMapper.updatePlanEnd(str,entry.getValue());
                }
            }

        }


            return new Result("200","success");
    }

    public Result InsertRunTask(String fileName) throws Exception {
        ganttTaskMapper.deleteRunTask();
        ganttTaskMapper.deleteRunGanttInfo();
//        String filePath = "D:\\CodePath\\test\\new_data.xlsx";
//        String filePath = "D:\\CodePath\\test\\副本2022年M+3实验需求Rev 03-整0318.xlsx";
        String filePath = fileName;
        Excel_Util.workbook = new XSSFWorkbook(filePath);
        List<RunGanttInfo> list_gantt = new LinkedList<>();
        HashMap<Integer,String> sheet_name = new HashMap<>();
        HashMap<String,String[]> Run_sheet = new HashMap<>();
        int col_input_amout_tmp = 0;
        String[] cycle = new String[5];
        String[] bank = new String[5];
        String[] table  = {"Array:0","EVEN:0","TPOT:0","EAC:0","Module:0"};
        String[] new_table = {"Array:1","EVEN:0","TPOT:0","EAC:0","Module:0"};
        DateFormat format=new SimpleDateFormat("yyyy/MM/dd");
        sheet_name.put(1,"ARRAY");
        sheet_name.put(2,"EVEN");
        sheet_name.put(3,"TPOT");
        sheet_name.put(4,"EAC");
        sheet_name.put(5,"MODULE");
        String sheetName = "";
        for(int i = 1; i <=5; i++){
            sheetName = sheet_name.get(i);
            int col_department = Excel_Util.readWantCol(filePath,sheetName,0,"实验部门");
            int col_customer = Excel_Util.readWantCol(filePath,sheetName,0,"客户名称");
            int col_pilot = Excel_Util.readWantCol(filePath,sheetName,0,"Pilot");
            int col_input = Excel_Util.readWantCol(filePath,sheetName,0,"投入时间");
            int col_output = Excel_Util.readWantCol(filePath,sheetName,0,"产出时间");
            int col_target = Excel_Util.readWantCol(filePath,sheetName,0,"实验目的");
            int col_input_amount = Excel_Util.readWantCol(filePath,sheetName,0,"投入数量");
            if(i == 4){
                col_input_amout_tmp = Excel_Util.readWantCol(filePath,sheetName,0,"集成度");
            }
            int col_output_amount = Excel_Util.readWantCol(filePath,sheetName,0,"产出数量");
            int col_number = Excel_Util.readWantCol(filePath,sheetName,0,"优先级");
            int col_prodect = Excel_Util.readWantCol(filePath,sheetName,0,"产品型号");
            int col_desc = Excel_Util.readWantCol(filePath,sheetName,0,"说明");
            int col_cycle = Excel_Util.readWantCol(filePath,sheetName,0,"cycle");
            cycle[i-1] = Excel_Util.readExcelData(filePath,sheetName,1,col_cycle);
            int col_bank = Excel_Util.readWantCol(filePath,sheetName,0,"bank");
            bank[i-1] = Excel_Util.readExcelData(filePath,sheetName,1,col_bank);
            int cur_rownum = Excel_Util.readrowNum(filePath,sheetName);
            for(int row = 1; row <= cur_rownum; row++){
                RunGanttInfo runGanttInfo = new RunGanttInfo();
                runGanttInfo.setTarget(Excel_Util.readExcelData(filePath,sheetName,row,col_target));
                runGanttInfo.setDepartment(Excel_Util.readExcelData(filePath,sheetName,row,col_department));
                runGanttInfo.setCustomer(Excel_Util.readExcelData(filePath,sheetName,row,col_customer));
                runGanttInfo.setDesc(Excel_Util.readExcelData(filePath,sheetName,row,col_desc));
                runGanttInfo.setCycle(Excel_Util.readExcelData(filePath,sheetName,row,col_cycle));
                runGanttInfo.setBank(Excel_Util.readExcelData(filePath,sheetName,row,col_bank));
                runGanttInfo.setDuration(1);
                if(!Excel_Util.readExcelData(filePath,sheetName,row,col_input).equals(""))
                    runGanttInfo.setInput_time(String.valueOf(format.format(org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.parseDouble(Excel_Util.readExcelData(filePath,sheetName,row,col_input))))));
                if(!Excel_Util.readExcelData(filePath,sheetName,row,col_output).equals(""))
                    runGanttInfo.setOutput_time(String.valueOf(format.format(org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.parseDouble(Excel_Util.readExcelData(filePath,sheetName,row,col_output))))));
                if(!Excel_Util.readExcelData(filePath,sheetName,row,col_number).equals(""))
                    runGanttInfo.setNumber(Double.parseDouble(Excel_Util.readExcelData(filePath,sheetName,row,col_number)));
                runGanttInfo.setProduct_number(Excel_Util.readExcelData(filePath,sheetName,row,col_prodect));
                runGanttInfo.setFactory_type(sheet_name.get(i));
                runGanttInfo.setPilot(Excel_Util.readExcelData(filePath,sheetName,row,col_pilot));
                if(!Excel_Util.readExcelData(filePath,sheetName,row,col_input_amount).equals(""))
                    runGanttInfo.setInput_amount(Double.parseDouble(Excel_Util.readExcelData(filePath,sheetName,row,col_input_amount)));
                if(!Excel_Util.readExcelData(filePath,sheetName,row,col_output_amount).equals("") && i != 4){
                    runGanttInfo.setOutput_amount(Double.parseDouble(Excel_Util.readExcelData(filePath,sheetName,row,col_output_amount)));
                    NumberFormat num = NumberFormat.getPercentInstance();
                    String rate_yield = num.format(runGanttInfo.getOutput_amount()/runGanttInfo.getInput_amount());
                    runGanttInfo.setYield(rate_yield);
                }
                else if(i == 4&&!Excel_Util.readExcelData(filePath,sheetName,row,col_output_amount).equals("")){
                    runGanttInfo.setOutput_amount(Double.parseDouble(Excel_Util.readExcelData(filePath,sheetName,row,col_output_amount)));
                    System.out.println(col_input_amout_tmp);
                    NumberFormat num = NumberFormat.getPercentInstance();
                    String rate_yield = num.format(runGanttInfo.getOutput_amount()/(runGanttInfo.getInput_amount()*Double.parseDouble(Excel_Util.readExcelData(filePath,sheetName,row,col_input_amout_tmp))));
                    runGanttInfo.setYield(rate_yield);
                }
                if(!Run_sheet.containsKey(runGanttInfo.getTarget())){
                    Run_sheet.put(runGanttInfo.getTarget(),table);
                    if(i == 1)
                        Run_sheet.put(runGanttInfo.getTarget(), new String[]{"Array:1", "EVEN:0", "TPOT:0", "EAC:0", "Module:0"});
                    else if(i == 2)
                        Run_sheet.put(runGanttInfo.getTarget(), new String[]{"Array:0", "EVEN:1", "TPOT:0", "EAC:0", "Module:0"});
                    else if(i == 3)
                        Run_sheet.put(runGanttInfo.getTarget(), new String[]{"Array:0", "EVEN:0", "TPOT:1", "EAC:0", "Module:0"});
                    else if(i == 4)
                        Run_sheet.put(runGanttInfo.getTarget(), new String[]{"Array:0", "EVEN:0", "TPOT:0", "EAC:1", "Module:0"});
                    else if(i == 5)
                        Run_sheet.put(runGanttInfo.getTarget(), new String[]{"Array:0", "EVEN:0", "TPOT:0", "EAC:0", "Module:1"});
//                    System.out.println(runGanttInfo.getTarget());
                }
                else {
                    new_table =  Run_sheet.get(runGanttInfo.getTarget());
                    if(i == 1)
                        new_table[0] = "Array:1";
                    else if(i == 2)
                        new_table[1] = "EVEN:1";
                    else if(i == 3)
                        new_table[2] = "TPOT:1";
                    else if(i == 4)
                        new_table[3] = "EAC:1";
                    else if(i == 5)
                        new_table[4] = "Module:1";
                    Run_sheet.put(runGanttInfo.getTarget(),new_table);
                }
//                System.out.println(runGanttInfo);
                list_gantt.add(runGanttInfo);
            }
//            System.out.println(sheet_name.get(i));
        }

        System.out.println(list_gantt.size());
//        HashMap<String,String[]> EVEN_sheet = readInputEven();

        list_gantt = infoUpdate(list_gantt,Run_sheet,cycle,bank);
        System.out.println(list_gantt.size());
        for(int i = 0; i < list_gantt.size();i++){
            if(!list_gantt.get(i).getTarget().equals(""))
                ganttTaskMapper.insertRunGanttInfo(list_gantt.get(i));
        }
        toGantt(list_gantt);
        updateByCapacity();
        return new Result("200","success");
    }
    public  HashMap<String,String[]> readInputEven(String fileName) throws Exception {
        String filePath = fileName;
        HashMap<String,String[]> EVEN_Sheet = new HashMap<>();
//        String filePath = "D:\\CodePath\\test\\副本2022年M+3实验需求Rev 03-整0318.xlsx";
        Excel_Util.workbook = new XSSFWorkbook(filePath);
        DateFormat format=new SimpleDateFormat("yyyy/MM/dd");
        int col_target = Excel_Util.readWantCol(filePath,"EVEN",0,"实验目的");
        int col_input = Excel_Util.readWantCol(filePath,"EVEN",0,"投入时间");
        int col_output = col_input+1;
        int col_in_amount = Excel_Util.readWantCol(filePath,"EVEN",0,"投入数量");
        int col_out_amount = col_in_amount+1;

        for(int row = 1; row <= Excel_Util.readrowNum(filePath,"EVEN");row++){
            String input = Excel_Util.DateToFormat(Excel_Util.readExcelData(filePath,"EVEN",row,col_input));
            String output = Excel_Util.DateToFormat(Excel_Util.readExcelData(filePath,"EVEN",row,col_output));
            String input_amount = Excel_Util.readExcelData(filePath,"EVEN",row,col_in_amount);
            String output_amount = Excel_Util.readExcelData(filePath,"EVEN",row,col_out_amount);
            EVEN_Sheet.put(Excel_Util.readExcelData(filePath,"EVEN",row,col_target), new String[]{input, output,input_amount,output_amount});
        }
//        System.out.println(EVEN_Sheet);
//                for (Map.Entry<String, String[]> entry : EVEN_Sheet.entrySet()) {
//            System.out.println("Key = " + entry.getKey());
//            for(int temp_i = 0; temp_i <entry.getValue().length;temp_i++){
//                System.out.println(entry.getValue()[temp_i]);
//            }
//        }
        return  EVEN_Sheet;
    }

    public  List<RunGanttInfo> infoUpdate(List<RunGanttInfo> list_gantt,HashMap<String,String[]> Run_sheet,String[] cycleTime, String[] bankTime) throws ParseException {
        HashMap<String,String> factory_count = new HashMap<>();
        Calendar rightNow = Calendar.getInstance();
        DateFormat format=new SimpleDateFormat("yyyy/MM/dd");
        HashMap<String,String> use_power = new HashMap<>();
        HashMap<String,String[]> EVEN_Sheet = new HashMap<>();
        System.out.println("!!!!!!!!!!!!!!!!!!!!!!!!!");
        for(int i = 0; i <cycleTime.length;i++){
            System.out.println("!!!!!!!!!!!!!");
            System.out.println(cycleTime[i]);
        }
        for(int i = 0; i < list_gantt.size();i++){
            if(list_gantt.get(i).getFactory_type().equals("EVEN")){
                EVEN_Sheet.put(list_gantt.get(i).getTarget(),new String[]{list_gantt.get(i).getInput_time(),list_gantt.get(i).getOutput_time(),String.valueOf(list_gantt.get(i).getInput_amount()),String.valueOf(list_gantt.get(i).getOutput_amount())});
            }
        }
        System.out.println("EVEN_sheet:");
        System.out.println(EVEN_Sheet.size());
        factory_count.put("Array","1");
        factory_count.put("EVEN","1");
        factory_count.put("TPOT","1");
        factory_count.put("EAC","1");
        factory_count.put("Module","1");
        for(int i = 0; i < ganttTaskMapper.getCapacity().size();i++){
            use_power.put(ganttTaskMapper.getCapacity().get(i).getFactory_type(),String.valueOf(ganttTaskMapper.getCapacity().get(i).getProduct_in_ability())+","+String.valueOf(ganttTaskMapper.getCapacity().get(i).getProduct_out_ability()));
        }
        for (Map.Entry<String, String[]> entry : Run_sheet.entrySet()) {
//            System.out.println("Key = " + entry.getKey());
            String even_flag = entry.getValue()[1].split(":")[1];
            String mod_flag = entry.getValue()[4].split(":")[1];
            for(int i = 0; i <list_gantt.size();i++){
                if(list_gantt.get(i).getTarget().equals(entry.getKey())&&!entry.getKey().equals("") &&EVEN_Sheet.get(entry.getKey()) != null){
                    if(even_flag.equals("1")){
//                        System.out.println(entry.getKey());
                        String even_input = EVEN_Sheet.get(entry.getKey())[0];
                        String even_output = EVEN_Sheet.get(entry.getKey())[1];
                        String even_in_amount = EVEN_Sheet.get(entry.getKey())[2];
                        String even_out_amount = EVEN_Sheet.get(entry.getKey())[3];
//                        list_gantt.get(i).setInput_amount(Double.parseDouble(even_in_amount));
//                        list_gantt.get(i).setOutput_amount(Double.parseDouble(even_out_amount));
//                        list_gantt.get(i).setInput_time(even_input);
//                        list_gantt.get(i).setOutput_time(even_output);

                        if(list_gantt.get(i).getFactory_type().equals("ARRAY")){
//                            System.out.println(list_gantt.get(i));
                            String inputTime = EVEN_Sheet.get(entry.getKey())[0];
                            String outputTime = EVEN_Sheet.get(entry.getKey())[1];
                            String cycleTime_Array = cycleTime[0];
                            String bankTime_Array  = bankTime[0];
                            System.out.println("ARRAY:"+cycleTime[0]);
                            Date date = format.parse(inputTime);
                            rightNow.setTime(date);
                            rightNow.add(Calendar.DAY_OF_YEAR, -1*Integer.parseInt(bankTime_Array));
                            String update_output = format.format(rightNow.getTime());
                            rightNow.setTime(rightNow.getTime());
                            rightNow.add(Calendar.DAY_OF_YEAR, -1*Integer.parseInt(cycleTime_Array));
                            String update_input = format.format(rightNow.getTime());
                            list_gantt.get(i).setOutput_time(update_output);
                            list_gantt.get(i).setInput_time(update_input);
                            System.out.println("---------------");
                            System.out.println(update_input);
                            System.out.println(update_output);
//                            System.out.println(list_gantt.get(i));
                        }
                        else if(list_gantt.get(i).getFactory_type().equals("TPOT")){
//                            System.out.println(list_gantt.get(i));
                            String inputTime = EVEN_Sheet.get(entry.getKey())[0];
                            String outputTime = EVEN_Sheet.get(entry.getKey())[1];
                            String cycleTime_TPOT = cycleTime[2];
                            String bankTime_to_TPOT  = bankTime[1];
                            Date date = format.parse(outputTime);
                            rightNow.setTime(date);
                            rightNow.add(Calendar.DAY_OF_YEAR, Integer.parseInt(bankTime_to_TPOT));
                            String update_input = format.format(rightNow.getTime());
                            rightNow.setTime(rightNow.getTime());
                            rightNow.add(Calendar.DAY_OF_YEAR, Integer.parseInt(cycleTime_TPOT));
                            String update_output = format.format(rightNow.getTime());
                            list_gantt.get(i).setOutput_time(update_output);
                            list_gantt.get(i).setInput_time(update_input);
                        }
                        else if(list_gantt.get(i).getFactory_type().equals("EAC")){
//                            System.out.println(list_gantt.get(i));
                            String inputTime = EVEN_Sheet.get(entry.getKey())[0];
                            String outputTime = EVEN_Sheet.get(entry.getKey())[1];
                            String cycleTime_ForTpot = cycleTime[2];
                            String bankTime_ForTpot  = bankTime[1];
                            String cycleTime_ForEAC = cycleTime[3];
                            String bankTime_ForEAC  = bankTime[2];
                            Date date = format.parse(outputTime);
                            rightNow.setTime(date);
                            rightNow.add(Calendar.DAY_OF_YEAR, Integer.parseInt(bankTime_ForTpot));
                            rightNow.add(Calendar.DAY_OF_YEAR, Integer.parseInt(cycleTime_ForTpot));
                            rightNow.add(Calendar.DAY_OF_YEAR, Integer.parseInt(bankTime_ForEAC));
                            String update_input = format.format(rightNow.getTime());
                            rightNow.setTime(rightNow.getTime());
                            rightNow.add(Calendar.DAY_OF_YEAR, Integer.parseInt(cycleTime_ForEAC));
                            String update_output = format.format(rightNow.getTime());
                            list_gantt.get(i).setOutput_time(update_output);
                            list_gantt.get(i).setInput_time(update_input);
//                            System.out.println(list_gantt.get(i));
                        }
                        else if(list_gantt.get(i).getFactory_type().equals("MODULE")){
//                            System.out.println(list_gantt.get(i));
                            String inputTime = EVEN_Sheet.get(entry.getKey())[0];
                            String outputTime = EVEN_Sheet.get(entry.getKey())[1];
                            String cycleTime_ForTpot = cycleTime[2];
                            String bankTime_ForTpot  = bankTime[1];
                            String cycleTime_ForEAC = cycleTime[3];
                            String bankTime_ForEAC  = bankTime[2];
                            String cycleTime_ForModule = cycleTime[4];
                            String bankTime_ForModule  = bankTime[3];
                            Date date = format.parse(outputTime);
                            rightNow.setTime(date);
                            rightNow.add(Calendar.DAY_OF_YEAR, Integer.parseInt(bankTime_ForTpot));
                            rightNow.add(Calendar.DAY_OF_YEAR, Integer.parseInt(cycleTime_ForTpot));
                            rightNow.add(Calendar.DAY_OF_YEAR, Integer.parseInt(bankTime_ForEAC));
                            rightNow.add(Calendar.DAY_OF_YEAR, Integer.parseInt(cycleTime_ForEAC));
                            rightNow.add(Calendar.DAY_OF_YEAR, Integer.parseInt(bankTime_ForModule));
                            String update_input = format.format(rightNow.getTime());
                            rightNow.setTime(rightNow.getTime());
                            rightNow.add(Calendar.DAY_OF_YEAR, Integer.parseInt(cycleTime_ForModule));
                            String update_output = format.format(rightNow.getTime());
                            list_gantt.get(i).setOutput_time(update_output);
                            list_gantt.get(i).setInput_time(update_input);
//                            System.out.println(list_gantt.get(i));
                        }
                    }
                    else if(even_flag.equals("0")){

                        if(mod_flag.equals("1")){
                            String inputTime = "";
                            String outputTime = "";
                            for(int tmp_i = 0; tmp_i < list_gantt.size(); tmp_i++){
                                if(list_gantt.get(tmp_i).getTarget().equals(list_gantt.get(i).getTarget())&&list_gantt.get(tmp_i).getFactory_type().equals("MODULE")){
                                    inputTime = list_gantt.get(tmp_i).getInput_time();
                                    outputTime = list_gantt.get(tmp_i).getOutput_time();
                                }
                            }
                            if(list_gantt.get(i).getFactory_type().equals("EAC")){
                                String cycleTime_EAC = cycleTime[3];
                                String bankTime_EAC  = bankTime[3];
                                Date date = format.parse(inputTime);
                                rightNow.setTime(date);
                                rightNow.add(Calendar.DAY_OF_YEAR, -1*Integer.parseInt(bankTime_EAC));
                                String update_output = format.format(rightNow.getTime());
                                rightNow.setTime(rightNow.getTime());
                                rightNow.add(Calendar.DAY_OF_YEAR, -1*Integer.parseInt(cycleTime_EAC));
                                String update_input = format.format(rightNow.getTime());
                                list_gantt.get(i).setOutput_time(update_output);
                                list_gantt.get(i).setInput_time(update_input);
                                System.out.println(list_gantt.get(i));
                            }

                            if(list_gantt.get(i).getFactory_type().equals("TPOT")){
                                String cycleTime_EAC = cycleTime[3];
                                String bankTime_EAC  = bankTime[3];
                                String cycleTime_Tpot = cycleTime[2];
                                String bankTime_Tpot  = bankTime[2];
                                Date date = format.parse(inputTime);
                                rightNow.setTime(date);
                                rightNow.add(Calendar.DAY_OF_YEAR, -1*Integer.parseInt(bankTime_EAC));
                                rightNow.add(Calendar.DAY_OF_YEAR, -1*Integer.parseInt(cycleTime_EAC));
                                rightNow.add(Calendar.DAY_OF_YEAR, -1*Integer.parseInt(bankTime_Tpot));
                                String update_output = format.format(rightNow.getTime());
                                rightNow.setTime(rightNow.getTime());
                                rightNow.add(Calendar.DAY_OF_YEAR, -1*Integer.parseInt(cycleTime_Tpot));
                                String update_input = format.format(rightNow.getTime());
                                list_gantt.get(i).setOutput_time(update_output);
                                list_gantt.get(i).setInput_time(update_input);
//                                System.out.println(list_gantt.get(i));
                            }

                            if(list_gantt.get(i).getFactory_type().equals("ARRAY")){
//                                System.out.println(list_gantt.get(i));
                                String cycleTime_EAC =cycleTime[3];
                                String bankTime_EAC  = bankTime[3];
                                String cycleTime_Tpot = cycleTime[2];
                                String bankTime_Tpot  = bankTime[2];
                                String cycleTime_EVEN = cycleTime[1];
                                String bankTime_EVEN  = bankTime[1];
                                String cycleTime_Array = cycleTime[0];
                                String bankTime_Array  = bankTime[1];
                                Date date = format.parse(inputTime);
                                rightNow.setTime(date);
                                rightNow.add(Calendar.DAY_OF_YEAR, -1*Integer.parseInt(bankTime_EAC));
                                rightNow.add(Calendar.DAY_OF_YEAR, -1*Integer.parseInt(cycleTime_EAC));
                                rightNow.add(Calendar.DAY_OF_YEAR, -1*Integer.parseInt(bankTime_Tpot));
                                rightNow.add(Calendar.DAY_OF_YEAR, -1*Integer.parseInt(cycleTime_Tpot));
                                rightNow.add(Calendar.DAY_OF_YEAR, -1*Integer.parseInt(bankTime_EVEN));
                                rightNow.add(Calendar.DAY_OF_YEAR, -1*Integer.parseInt(cycleTime_EVEN));
                                rightNow.add(Calendar.DAY_OF_YEAR, -1*Integer.parseInt(bankTime_Array));
                                String update_output = format.format(rightNow.getTime());
                                rightNow.setTime(rightNow.getTime());
                                rightNow.add(Calendar.DAY_OF_YEAR, -1*Integer.parseInt(cycleTime_Array));
                                String update_input = format.format(rightNow.getTime());
                                list_gantt.get(i).setOutput_time(update_output);
                                list_gantt.get(i).setInput_time(update_input);
//                                System.out.println(list_gantt.get(i));
                            }

                        }
                    }
                }
//                else if(list_gantt.get(i).getTarget().equals(entry.getKey())&&!entry.getKey().equals("") &&EVEN_Sheet.get(entry.getKey()) == null){
//                    list_gantt.remove(i);
//                    System.out.println(list_gantt.get(i).getTarget());
//                }
            }
//            System.out.println("----------------");
        }
        for(int i = 0; i <list_gantt.size();i++){
            System.out.println(list_gantt.get(i));
        }
        return list_gantt;
    }

    public  void toGantt(List<RunGanttInfo> list_gantt){
        ganttTaskMapper.deleteRunTask();
        Set<String> father_target = new HashSet();
        for(int i = 0; i <list_gantt.size();i++){
            RunGanttTask father = new RunGanttTask();
            RunGanttTask parent = new RunGanttTask();
            RunGanttTask runGantttask_input = new RunGanttTask();
            RunGanttTask runGantttask_output = new RunGanttTask();
            if(!father_target.contains(list_gantt.get(i).getTarget())&&!list_gantt.get(i).getTarget().equals("")){
                father_target.add(list_gantt.get(i).getTarget());
                father.setColor("rgba(0,0,0,0)");
                father.setPilot(list_gantt.get(i).getPilot());
                father.setOpen("true");
                father.setId(list_gantt.get(i).getTarget()+"father");
                father.setText(list_gantt.get(i).getTarget());
                System.out.println(father);
                ganttTaskMapper.insertTask_run(father);
            }
            parent.setColor("rgba(0,0,0,0)");
            parent.setParent(list_gantt.get(i).getTarget()+"father");
            parent.setText(list_gantt.get(i).getTarget());
            parent.setOpen("true");
            parent.setRender("split");

            parent.setPilot(list_gantt.get(i).getPilot());
            parent.setId(list_gantt.get(i).getTarget()+"_"+list_gantt.get(i).getFactory_type());
            parent.setFactory_type(list_gantt.get(i).getFactory_type());
            System.out.println(parent.getFactory_type());
            if(parent.getFactory_type().equals("ARRAY")){
                parent.setNumber(1);
            }
            if(parent.getFactory_type().equals("EVEN")){
                parent.setNumber(2);
            }
            if(parent.getFactory_type().equals("TPOT")){
                parent.setNumber(3);
            }
            if(parent.getFactory_type().equals("EAC")){
                parent.setNumber(4);
            }
            if(parent.getFactory_type().equals("MODULE")){
                parent.setNumber(5);
            }
            System.out.println(parent);
            ganttTaskMapper.insertTask_run(parent);
            runGantttask_input.setParent(parent.getId());
            runGantttask_input.setId(parent.getId()+"投入");
            runGantttask_input.setPilot(parent.getPilot());
            runGantttask_input.setText(runGantttask_input.getId());
            runGantttask_input.setFactory_type(parent.getFactory_type());
            runGantttask_input.setColor("rgba(255,165,0,0.5)");
            runGantttask_input.setStart_date(list_gantt.get(i).getInput_time());
            runGantttask_input.setUse_amount(list_gantt.get(i).getInput_amount());
            runGantttask_input.setNumber(parent.getNumber());
            runGantttask_input.setDuration(1);
            System.out.println(runGantttask_input);
            ganttTaskMapper.insertTask_run(runGantttask_input);
            runGantttask_output.setDuration(1);
            runGantttask_output.setParent(parent.getId());
            runGantttask_output.setPilot(parent.getPilot());
            runGantttask_output.setParent(parent.getId());
            runGantttask_output.setId(parent.getId()+"产出");
            runGantttask_output.setFactory_type(parent.getFactory_type());
            runGantttask_output.setText(runGantttask_output.getId());
            runGantttask_output.setColor("rgba(192,192,192,0.5)");
            runGantttask_output.setStart_date(list_gantt.get(i).getOutput_time());
            runGantttask_output.setUse_amount(list_gantt.get(i).getOutput_amount());
            System.out.println(runGantttask_output);
            ganttTaskMapper.insertTask_run(runGantttask_output);
        }

    }

    public void runGanttdownload() throws IOException, ParseException {
        List<RunGanttInfo> ganttInfos = ganttTaskMapper.getRunGanttInfo();
        List<GanttCapacity> ganttCapacities = ganttTaskMapper.getCapacity();
        HashMap<Integer,GanttCapacity> Map_CapaCity = new HashMap<>();
        for(int i = 0 ; i < 5; i++){
            Map_CapaCity.put(i,ganttCapacities.get(i));
        }
        writePlanExcel(ganttInfos,Map_CapaCity);
        System.out.println(ganttInfos);
        System.out.println(Map_CapaCity);
    }

    public void writePlanExcel(List<RunGanttInfo> list_gantt,HashMap<Integer,GanttCapacity> Map_CapaCity) throws IOException, ParseException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet_array = workbook.createSheet("Array计划");
        Sheet sheet_even = workbook.createSheet("EVEN计划");
        Sheet sheet_tpot = workbook.createSheet("TPOT计划");
        Sheet sheet_eac = workbook.createSheet("EAC计划");
        Sheet sheet_module = workbook.createSheet("MODULE计划");
        DateFormat df = new SimpleDateFormat("yyyy/MM/dd");
//        List<GanttCapacity> Map_CapaCity = new LinkedList<>();
        HashMap<Integer,Sheet> map_sheet = new HashMap();
        HashMap<Integer,String> factory_mark = new HashMap();
        HashMap<Integer,List<RunGanttInfo>> factory_gantt = new HashMap();

        List<RunGanttInfo> list_gantt_array = new LinkedList<>();
        List<RunGanttInfo> list_gantt_even = new LinkedList<>();
        List<RunGanttInfo> list_gantt_tpot = new LinkedList<>();
        List<RunGanttInfo> list_gantt_eac = new LinkedList<>();
        List<RunGanttInfo> list_gantt_module = new LinkedList<>();

        for(int i = 0; i <list_gantt.size();i++){
            if(list_gantt.get(i).getFactory_type().equals("ARRAY"))
                list_gantt_array.add(list_gantt.get(i));
            if(list_gantt.get(i).getFactory_type().equals("EVEN"))
                list_gantt_even.add(list_gantt.get(i));
            if(list_gantt.get(i).getFactory_type().equals("TPOT"))
                list_gantt_tpot.add(list_gantt.get(i));
            if(list_gantt.get(i).getFactory_type().equals("EAC"))
                list_gantt_eac.add(list_gantt.get(i));
            if(list_gantt.get(i).getFactory_type().equals("MODULE"))
                list_gantt_module.add(list_gantt.get(i));
        }
        factory_gantt.put(0,list_gantt_array);
        factory_gantt.put(1,list_gantt_even);
        factory_gantt.put(2,list_gantt_tpot);
        factory_gantt.put(3,list_gantt_eac);
        factory_gantt.put(4,list_gantt_module);
        Calendar calendar = new GregorianCalendar();
        map_sheet.put(0,sheet_array);
        map_sheet.put(1,sheet_even);
        map_sheet.put(2,sheet_tpot);
        map_sheet.put(3,sheet_eac);
        map_sheet.put(4,sheet_module);
        factory_mark.put(0,"ARRAY");
        factory_mark.put(1,"EVEN");
        factory_mark.put(2,"TPOT");
        factory_mark.put(3,"EAC");
        factory_mark.put(4,"MODULE");
        List<String> head_list = new LinkedList<>();
        List<Date> date_list = new LinkedList<>();
        head_list.add("厂别");
        head_list.add("Pilot");
        head_list.add("实验部门");
        head_list.add("客户名称");
        head_list.add("产品型号");
        head_list.add("说明");
        head_list.add("实验目的");
        head_list.add("优先级");
        head_list.add("投入时间");
        head_list.add("产出时间");
        head_list.add("cycle时间");
        head_list.add("bank时间");
        head_list.add("投入数量");
        head_list.add("产出数量");
        head_list.add("预估良率");
        head_list.add("IN/OUT");

        Date min_input =new Date( Long.MAX_VALUE);
        Date max_ouput = new Date(0);
        Date cur_input_time = new Date();
        Date cur_output_time = new Date();


        for(int i = 0; i < list_gantt.size();i++){
            if(list_gantt.get(i).getInput_time() != null && list_gantt.get(i).getOutput_time() != null){
                cur_input_time = df.parse(list_gantt.get(i).getInput_time());
                cur_output_time = df.parse(list_gantt.get(i).getOutput_time());
                if(cur_input_time.getTime() - min_input.getTime() < 0){
                    min_input = cur_input_time;
                }
                if(cur_output_time.getTime() - max_ouput.getTime() > 0){
                    max_ouput = cur_output_time;
                }
            }
        }
        Date tmp_date = min_input;
        while(tmp_date.getTime() <= max_ouput.getTime()){
//            System.out.println(df.format(tmp_date));
            date_list.add(tmp_date);
            calendar.setTime(tmp_date);
            calendar.add(calendar.DATE,1);
            tmp_date = calendar.getTime();
        }

        for(int i = 0; i < 5; i++){
            XSSFRow row = (XSSFRow) map_sheet.get(i).createRow(0);
            XSSFCell cell;
            for(int tmp_i = 0; tmp_i < head_list.size(); tmp_i++){
                cell = row.createCell(tmp_i);
                cell.setCellValue(head_list.get(tmp_i));
            }
            for(int tmp_i = 0; tmp_i < date_list.size(); tmp_i++){
                cell = row.createCell(tmp_i+16);
                cell.setCellValue(date_list.get(tmp_i));
            }
        }

        for(int index = 0; index < 5; index++){
            List<RunGanttInfo> cur_gantt = factory_gantt.get(index);
            for(int i = 0, tmp_i = 0; i < cur_gantt.size(); i++,tmp_i = tmp_i + 2){
                XSSFRow row_in = (XSSFRow) map_sheet.get(index).createRow(tmp_i+1 );
                XSSFRow row_out = (XSSFRow) map_sheet.get(index).createRow(tmp_i+2 );
                XSSFCell cell_in;
                XSSFCell cell_out;
                for(int cellnum = 0; cellnum < 17; cellnum++){
                    cell_in = row_in.createCell(cellnum);
                    cell_out = row_out.createCell(cellnum);
                    if(cellnum == 0&&factory_mark.get(index).equals(cur_gantt.get(i).getFactory_type())&&!cur_gantt.get(i).getTarget().equals("")){
                        cell_in.setCellValue(cur_gantt.get(i).getFactory_type());
                        cell_out.setCellValue(cur_gantt.get(i).getFactory_type());
                    }
                    if(cellnum == 1&&factory_mark.get(index).equals(cur_gantt.get(i).getFactory_type())&&!cur_gantt.get(i).getTarget().equals("")){
                        cell_in.setCellValue(cur_gantt.get(i).getPilot());
                        cell_out.setCellValue(cur_gantt.get(i).getPilot());
                    }
                    if(cellnum == 2&&factory_mark.get(index).equals(cur_gantt.get(i).getFactory_type())&&!cur_gantt.get(i).getTarget().equals("")){
                        cell_in.setCellValue(cur_gantt.get(i).getDepartment());
                        cell_out.setCellValue(cur_gantt.get(i).getDepartment());
                    }
                    if(cellnum == 3&&factory_mark.get(index).equals(cur_gantt.get(i).getFactory_type())&&!cur_gantt.get(i).getTarget().equals("")){
                        cell_in.setCellValue(cur_gantt.get(i).getCustomer());
                        cell_out.setCellValue(cur_gantt.get(i).getCustomer());
                    }
                    if(cellnum == 4&&factory_mark.get(index).equals(cur_gantt.get(i).getFactory_type())&&!cur_gantt.get(i).getTarget().equals("")){
                        cell_in.setCellValue(cur_gantt.get(i).getProduct_number());
                        cell_out.setCellValue(cur_gantt.get(i).getProduct_number());
                    }
                    if(cellnum == 5&&factory_mark.get(index).equals(cur_gantt.get(i).getFactory_type())&&!cur_gantt.get(i).getTarget().equals("")){
                        cell_in.setCellValue(cur_gantt.get(i).getDesc());
                        cell_out.setCellValue(cur_gantt.get(i).getDesc());
                    }
                    if(cellnum == 6&&factory_mark.get(index).equals(cur_gantt.get(i).getFactory_type())&&!cur_gantt.get(i).getTarget().equals("")){
                        cell_in.setCellValue(cur_gantt.get(i).getTarget());
                        cell_out.setCellValue(cur_gantt.get(i).getTarget());
                    }
                    if(cellnum == 7&&factory_mark.get(index).equals(cur_gantt.get(i).getFactory_type())&&!cur_gantt.get(i).getTarget().equals("")){
                        if(cur_gantt.get(i).getNumber() != null){
                            cell_in.setCellValue(cur_gantt.get(i).getNumber());
                            cell_out.setCellValue(cur_gantt.get(i).getNumber());
                        }
                        else{
                            cell_in.setCellValue("");
                            cell_out.setCellValue("");
                        }
                    }
                    if(cellnum == 8&&factory_mark.get(index).equals(cur_gantt.get(i).getFactory_type())&&!cur_gantt.get(i).getTarget().equals("")){
                        cell_in.setCellValue(cur_gantt.get(i).getInput_time());
                        cell_out.setCellValue(cur_gantt.get(i).getInput_time());
                    }
                    if(cellnum == 9&&factory_mark.get(index).equals(cur_gantt.get(i).getFactory_type())&&!cur_gantt.get(i).getTarget().equals("")){
                        cell_in.setCellValue(cur_gantt.get(i).getOutput_time());
                        cell_out.setCellValue(cur_gantt.get(i).getOutput_time());
                    }

                    if(cellnum == 12&&factory_mark.get(index).equals(cur_gantt.get(i).getFactory_type())&&!cur_gantt.get(i).getTarget().equals("")){
                        cell_in.setCellValue(cur_gantt.get(i).getInput_amount());
                        cell_out.setCellValue(cur_gantt.get(i).getInput_amount());
                    }
                    if(cellnum == 13&&factory_mark.get(index).equals(cur_gantt.get(i).getFactory_type())&&!cur_gantt.get(i).getTarget().equals("")){
                        cell_in.setCellValue(cur_gantt.get(i).getOutput_amount());
                        cell_out.setCellValue(cur_gantt.get(i).getOutput_amount());
                    }
                    if(cellnum == 14&&factory_mark.get(index).equals(cur_gantt.get(i).getFactory_type())&&!cur_gantt.get(i).getTarget().equals("")){
                        cell_in.setCellValue(cur_gantt.get(i).getYield());
                        cell_out.setCellValue(cur_gantt.get(i).getYield());
                    }
                    if(cellnum == 15&&factory_mark.get(index).equals(cur_gantt.get(i).getFactory_type())&&!cur_gantt.get(i).getTarget().equals("")){
                        cell_in.setCellValue("投入");
                        cell_out.setCellValue("产出");
                    }
                }
                for(int cellnum = 17; cellnum < map_sheet.get(index).getRow(0).getPhysicalNumberOfCells();cellnum++){
//                    System.out.println();

                    Date cellValue = map_sheet.get(index).getRow(0).getCell(cellnum).getDateCellValue();
                    if(df.format(cellValue).equals(cur_gantt.get(i).getInput_time()))
                    {
                        Double tmp_amount = cur_gantt.get(i).getInput_amount();
                        for(int j = 0; j <Math.ceil(cur_gantt.get(i).getInput_amount()/Map_CapaCity.get(index).getProduct_in_ability());j++){
                            cell_in = row_in.createCell(cellnum);
                            if(tmp_amount - Map_CapaCity.get(index).getProduct_in_ability() < 0){
                                cell_in = row_in.createCell(cellnum);
                                cell_in.setCellValue(tmp_amount);
                            }
                            else if(tmp_amount - Map_CapaCity.get(index).getProduct_in_ability() >0){
                                tmp_amount -= Map_CapaCity.get(index).getProduct_in_ability();
                                cell_in.setCellValue(Map_CapaCity.get(index).getProduct_in_ability());
                                cellnum++;
                            }
                        }
                    }
                    if(df.format(cellValue).equals(cur_gantt.get(i).getOutput_time()))
                    {
                        Double tmp_amount = cur_gantt.get(i).getOutput_amount();
                        for(int j = 0; j <Math.ceil(cur_gantt.get(i).getOutput_amount()/Map_CapaCity.get(index).getProduct_out_ability());j++){

                            cell_out = row_out.createCell(cellnum);
                            if(tmp_amount - Map_CapaCity.get(index).getProduct_out_ability() < 0){
                                cell_out = row_out.createCell(cellnum);
                                cell_out.setCellValue(tmp_amount);
                            }
                            else if(tmp_amount - Map_CapaCity.get(index).getProduct_out_ability() >0){
                                tmp_amount -= Map_CapaCity.get(index).getProduct_out_ability();
                                cell_out.setCellValue(Map_CapaCity.get(index).getProduct_out_ability());
                                cellnum++;
                            }
                        }
                    }
                }
            }
        }

        String fileName="运营部生成计划" + ".xlsx";
        File desktopDir= FileSystemView.getFileSystemView().getHomeDirectory();//获取桌面的目录
        String desktopPath=desktopDir.getAbsolutePath();//获取桌面的绝对路径
        String filePath="D:\\yunying\\upload\\excel\\test"+"\\"+fileName;
        FileOutputStream out=new FileOutputStream(filePath);
        workbook.write(out);
    }

    public void updateByCapacity() throws ParseException{
        ganttTaskMapper.deleteRunTask();
        List<RunGanttInfo> list_gantt = ganttTaskMapper.getRunGanttInfo();
        List<RunGanttInfo> list_tmp = new LinkedList<>();
        HashMap<String,String> use_power = new HashMap<>();
        HashMap<String,Double> map_time = new HashMap<>();
        HashMap<String,String[]> Run_sheet = new HashMap<>();
        Calendar calendar = Calendar.getInstance();
        Boolean flag = true;
        HashMap<Integer,String> facotry_name = new HashMap<>();
        facotry_name.put(0,"ARRAY");
        facotry_name.put(1,"TPOT");
        facotry_name.put(2,"EAC");
        facotry_name.put(3,"MODULE");
        String[] table  = {"Array:0","EVEN:0","TPOT:0","EAC:0","Module:0"};
        String[] new_table = {"Array:1","EVEN:0","TPOT:0","EAC:0","Module:0"};
        DateFormat format=new SimpleDateFormat("yyyy/MM/dd");

        for(int i = 0; i < ganttTaskMapper.getCapacity().size();i++){
            use_power.put(ganttTaskMapper.getCapacity().get(i).getFactory_type(),String.valueOf(ganttTaskMapper.getCapacity().get(i).getProduct_in_ability())+","+String.valueOf(ganttTaskMapper.getCapacity().get(i).getProduct_out_ability()));
        }
        for(int i = 0; i < list_gantt.size();i++){
            if(!Run_sheet.containsKey(list_gantt.get(i).getTarget())){
                Run_sheet.put(list_gantt.get(i).getTarget(),table);
                if(list_gantt.get(i).getFactory_type().equals("ARRAY"))
                    Run_sheet.put(list_gantt.get(i).getTarget(), new String[]{"Array:1", "EVEN:0", "TPOT:0", "EAC:0", "Module:0"});
                else if(list_gantt.get(i).getFactory_type().equals("EVEN"))
                    Run_sheet.put(list_gantt.get(i).getTarget(), new String[]{"Array:0", "EVEN:1", "TPOT:0", "EAC:0", "Module:0"});
                else if(list_gantt.get(i).getFactory_type().equals("TPOT"))
                    Run_sheet.put(list_gantt.get(i).getTarget(), new String[]{"Array:0", "EVEN:0", "TPOT:1", "EAC:0", "Module:0"});
                else if(list_gantt.get(i).getFactory_type().equals("EAC"))
                    Run_sheet.put(list_gantt.get(i).getTarget(), new String[]{"Array:0", "EVEN:0", "TPOT:0", "EAC:1", "Module:0"});
                else if(list_gantt.get(i).getFactory_type().equals("MODULE"))
                    Run_sheet.put(list_gantt.get(i).getTarget(), new String[]{"Array:0", "EVEN:0", "TPOT:0", "EAC:0", "Module:1"});
//                    System.out.println(runGanttInfo.getTarget());
            }
            else {
                new_table =  Run_sheet.get(list_gantt.get(i).getTarget());
                if(list_gantt.get(i).getFactory_type().equals("ARRAY"))
                    new_table[0] = "Array:1";
                else if(list_gantt.get(i).getFactory_type().equals("EVEN"))
                    new_table[1] = "EVEN:1";
                else if(list_gantt.get(i).getFactory_type().equals("TPOT"))
                    new_table[2] = "TPOT:1";
                else if(list_gantt.get(i).getFactory_type().equals("EAC"))
                    new_table[3] = "EAC:1";
                else if(list_gantt.get(i).getFactory_type().equals("MODULE"))
                    new_table[4] = "Module:1";
                Run_sheet.put(list_gantt.get(i).getTarget(),new_table);
            }
        }
        for(int index = 0; index < 1; index++) {
            list_tmp.clear();
            for (int i = 0; i < list_gantt.size(); i++) {
                if (list_gantt.get(i).getFactory_type().equals(facotry_name.get(index))) {
                    list_tmp.add(list_gantt.get(i));
                    map_time.put(list_gantt.get(i).getInput_time(), Double.parseDouble(use_power.get(facotry_name.get(index)).split(",")[0]));
                }
            }

            for (int i = 0; i < list_tmp.size(); i++) {
                System.out.println(list_tmp.get(i).getFactory_type()+":"+list_tmp.get(i).getInput_time() + ":" + list_tmp.get(i).getInput_amount() + ":" + list_tmp.get(i).getDuration());
            }

            Collections.sort(list_tmp);
            int count = 0;
            for (int i = 0; i < list_tmp.size(); i++) {
                flag = true;
                double cap = 0; //
                double tmp_cap = list_tmp.get(i).getInput_amount();
                int tmp_count = 0;
                Double over = -1.0;
                Double next_avli = 0.0;
                if(index == 0) {
                    while (flag) {
                        count++;
                        Double tmp_input_amount = list_tmp.get(i).getInput_amount();
                        tmp_input_amount = tmp_input_amount - cap;
                        System.out.println(tmp_input_amount);
                        double t = map_time.get(list_tmp.get(i).getInput_time());
                        if (map_time.get(list_tmp.get(i).getInput_time()) - tmp_input_amount >= 0 || (over <= Double.parseDouble(use_power.get("ARRAY").split(",")[0]) && over >= 0)) {
//                    map_time.put(list_tmp.get(i).getInput_time(), map_time.get(list_tmp.get(i).getInput_time()) - list_tmp.get(i).getInput_amount());
                            Double tmp_over = map_time.get(list_tmp.get(i).getInput_time()) - list_tmp.get(i).getInput_amount();
                            if (tmp_over < 0)
                                map_time.put(list_tmp.get(i).getInput_time(), over);
                            else
                                map_time.put(list_tmp.get(i).getInput_time(), map_time.get(list_tmp.get(i).getInput_time()) - list_tmp.get(i).getInput_amount());

                            flag = false;
                        }
                        else {

                            if (list_tmp.get(i).getInput_time().equals("2022/05/15")) {
                                System.out.println("1");
                            }

//                    tmp_count++;
//                    if(tmp_count > 10){
//                        System.out.println("&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&");
//                        System.out.println(list_tmp.get(i));
//                        break;
//                    }

//                    if(list_tmp.get(i).getInput_amount() == 300.0){
//                        System.out.println("1");
//                    }
                            System.out.println("2222222222");
                            System.out.println(list_tmp.get(i).getInput_time());
                            calendar.setTime(format.parse(list_tmp.get(i).getInput_time()));
                            double test = map_time.get(list_tmp.get(i).getInput_time());
                            calendar.add(calendar.DATE, -1);
//                    cap += tmp_cap - map_time.get(list_tmp.get(i).getInput_time());
                            cap += test;
                            tmp_cap -= map_time.get(list_tmp.get(i).getInput_time());
//                    tmp_input_amount = tmp_input_amount - cap;
                            if (map_time.containsKey(format.format(calendar.getTime()))) {
                                System.out.println(format.format(calendar.getTime()) + ":" + map_time.get(format.format(calendar.getTime())));
                                next_avli = map_time.get(format.format(calendar.getTime()));
                            } else
                                next_avli = Double.parseDouble(use_power.get("ARRAY").split(",")[0]);

                            if (test == 0.0) {
                                over = next_avli + over;
                            } else
                                over = next_avli - (list_tmp.get(i).getInput_amount() - map_time.get(list_tmp.get(i).getInput_time()));
                            if (over >= 0) {
//                        if(map_time.containsKey(format.format(calendar.getTime())))
//                            map_time.put(format.format(calendar.getTime()),map_time.get(format.format(calendar.getTime()))- (list_tmp.get(i).getInput_amount() - map_time.get(list_tmp.get(i).getInput_time())));
//                        else
                                System.out.println(format.format(calendar.getTime()) + ":" + over);
                                map_time.put(format.format(calendar.getTime()), over);

                            } else {
                                map_time.put(format.format(calendar.getTime()), 0.0);
                            }
                            map_time.put(list_tmp.get(i).getInput_time(), 0.0);
                            list_tmp.get(i).setInput_time(format.format(calendar.getTime()));
                            String time = list_tmp.get(i).getInput_time();
                            list_tmp.get(i).setDuration(list_tmp.get(i).getDuration() + 1);
                            System.out.println(list_tmp.get(i).getInput_time() + ":" + list_tmp.get(i).getInput_amount() + ":" + list_tmp.get(i).getDuration());
                        }
                    }
                }
                else {
                    while (flag){
//                        tmp_input_amount 该实验目前剩余多少需要投入
//                        cap 迭代累加补满当天所需的量
//                        over 下一天扣除掉当天剩余的投入量之后的剩余值
                        Double tmp_input_amount = list_tmp.get(i).getInput_amount();
                        tmp_input_amount = tmp_input_amount - cap;
                        System.out.println(tmp_input_amount);
                        double t = map_time.get(list_tmp.get(i).getInput_time()); //获取对应时间还可容纳的量
                        if (t - tmp_input_amount >= 0 || (over <= Double.parseDouble(use_power.get(facotry_name.get(index)).split(",")[0]) && over >= 0)) {
//                    map_time.put(list_tmp.get(i).getInput_time(), map_time.get(list_tmp.get(i).getInput_time()) - list_tmp.get(i).getInput_amount());
                            Double tmp_over = map_time.get(list_tmp.get(i).getInput_time()) - list_tmp.get(i).getInput_amount();
                            if (tmp_over < 0)
                                map_time.put(list_tmp.get(i).getInput_time(), over);
                            else
                                map_time.put(list_tmp.get(i).getInput_time(), map_time.get(list_tmp.get(i).getInput_time()) - list_tmp.get(i).getInput_amount());
                            flag = false;
                        }
                        else {
                            calendar.setTime(format.parse(list_tmp.get(i).getInput_time()));
                            calendar.add(calendar.DATE, -1);
                            Double today_avli = map_time.get(list_tmp.get(i).getInput_time());
                            cap += today_avli;
                            if (map_time.containsKey(format.format(calendar.getTime()))) {
                                System.out.println(format.format(calendar.getTime()) + ":" + map_time.get(format.format(calendar.getTime())));
                                next_avli = map_time.get(format.format(calendar.getTime()));
                            } else
                                next_avli = Double.parseDouble(use_power.get(facotry_name.get(index)).split(",")[0]);

                            if (today_avli == 0.0 && over < 0) {
                                over = next_avli + over;
                            } else
                                over = next_avli - (list_tmp.get(i).getInput_amount() - map_time.get(list_tmp.get(i).getInput_time()));
                            if (over >= 0) {
//                                System.out.println(format.format(calendar.getTime()) + ":" + over);
                                map_time.put(format.format(calendar.getTime()), over);

                            } else {
                                map_time.put(format.format(calendar.getTime()), 0.0);
                            }
                            map_time.put(list_tmp.get(i).getInput_time(), 0.0);
                            list_tmp.get(i).setInput_time(format.format(calendar.getTime()));
                            String time = list_tmp.get(i).getInput_time();
                            list_tmp.get(i).setDuration(list_tmp.get(i).getDuration() + 1);
                            System.out.println(list_tmp.get(i).getInput_time() + ":" + list_tmp.get(i).getInput_amount() + ":" + list_tmp.get(i).getDuration());
                        }
                    }
//                    for (int d = 0; d < list_tmp.size(); d++) {
//                        System.out.println(list_tmp.get(d).getFactory_type()+list_tmp.get(d).getInput_time() + ":" + list_gantt.get(d).getInput_amount() + ":" + list_gantt.get(d).getDuration());
//                    }
                }

            }

            for (Map.Entry<String, Double> entry : map_time.entrySet()) {
                System.out.println(entry.getKey() + ":" + entry.getValue());
            }
            Collections.sort(list_tmp);
            for (int i = 0; i < list_tmp.size(); i++) {
                System.out.println(list_tmp.get(i).getFactory_type()+":"+list_tmp.get(i).getInput_time() + ":" + list_tmp.get(i).getInput_amount() + ":" + list_tmp.get(i).getDuration());
                calendar.setTime(format.parse(list_tmp.get(i).getInput_time()));
                calendar.add(calendar.DATE, Integer.parseInt(list_tmp.get(i).getCycle()));
                list_tmp.get(i).setOutput_time(format.format(calendar.getTime()));
            }
            infoToGantt(list_tmp);

        }
    }
    public void infoToGantt(List<RunGanttInfo> list_gantt){

        Set<String> father_target = new HashSet();
        for(int i = 0; i <list_gantt.size();i++){
            RunGanttTask father = new RunGanttTask();
            RunGanttTask parent = new RunGanttTask();
            RunGanttTask runGantttask_input = new RunGanttTask();
            RunGanttTask runGantttask_output = new RunGanttTask();
            if(!father_target.contains(list_gantt.get(i).getTarget())&&!list_gantt.get(i).getTarget().equals("")){
                father_target.add(list_gantt.get(i).getTarget());
                father.setColor("rgba(0,0,0,0)");
                father.setPilot(list_gantt.get(i).getPilot());
                father.setOpen("true");
                father.setId(list_gantt.get(i).getTarget()+"father");
                father.setText(list_gantt.get(i).getTarget());
                System.out.println(father);
                ganttTaskMapper.insertTask_run(father);
            }
            parent.setColor("rgba(0,0,0,0)");
            parent.setParent(list_gantt.get(i).getTarget()+"father");
            parent.setText(list_gantt.get(i).getTarget());
            parent.setOpen("true");
            parent.setRender("split");

            parent.setPilot(list_gantt.get(i).getPilot());
            parent.setId(list_gantt.get(i).getTarget()+"_"+list_gantt.get(i).getFactory_type());
            parent.setFactory_type(list_gantt.get(i).getFactory_type());
            System.out.println(parent.getFactory_type());
            if(parent.getFactory_type().equals("ARRAY")){
                parent.setNumber(1);
            }
            if(parent.getFactory_type().equals("EVEN")){
                parent.setNumber(2);
            }
            if(parent.getFactory_type().equals("TPOT")){
                parent.setNumber(3);
            }
            if(parent.getFactory_type().equals("EAC")){
                parent.setNumber(4);
            }
            if(parent.getFactory_type().equals("MODULE")){
                parent.setNumber(5);
            }
            System.out.println(parent);
            ganttTaskMapper.insertTask_run(parent);
            runGantttask_input.setParent(parent.getId());
            runGantttask_input.setId(parent.getId()+"投入");
            runGantttask_input.setPilot(parent.getPilot());
            runGantttask_input.setText(runGantttask_input.getId());
            runGantttask_input.setFactory_type(parent.getFactory_type());
            runGantttask_input.setColor("rgba(255,165,0,0.5)");
            runGantttask_input.setStart_date(list_gantt.get(i).getInput_time());
            runGantttask_input.setUse_amount(list_gantt.get(i).getInput_amount());
            runGantttask_input.setNumber(parent.getNumber());
            runGantttask_input.setDuration(list_gantt.get(i).getDuration()+1);
            System.out.println(runGantttask_input);
            ganttTaskMapper.insertTask_run(runGantttask_input);
            runGantttask_output.setDuration(1);
            runGantttask_output.setParent(parent.getId());
            runGantttask_output.setPilot(parent.getPilot());
            runGantttask_output.setParent(parent.getId());
            runGantttask_output.setId(parent.getId()+"产出");
            runGantttask_output.setFactory_type(parent.getFactory_type());
            runGantttask_output.setText(runGantttask_output.getId());
            runGantttask_output.setColor("rgba(192,192,192,0.5)");
            runGantttask_output.setStart_date(list_gantt.get(i).getOutput_time());
            runGantttask_output.setUse_amount(list_gantt.get(i).getOutput_amount());
            System.out.println(runGantttask_output);
            ganttTaskMapper.insertTask_run(runGantttask_output);
        }
    }
//    public void updateByCapacity() throws ParseException {
//        ganttTaskMapper.deleteRunTask();
//        List<RunGanttInfo> list_gantt = ganttTaskMapper.getRunGanttInfo();
//        List<RunGanttInfo> list_gantt_array = new LinkedList<>();
//        HashMap<String,String> use_power = new HashMap<>();
//        HashMap<String,Double> over_power_date = new HashMap<>();
//        HashMap<String,Double> map_time_array = new HashMap<>();
//        HashMap<String,String[]> Run_sheet = new HashMap<>();
//        Calendar calendar = Calendar.getInstance();
//        DateFormat format=new SimpleDateFormat("yyyy/MM/dd");
//        boolean flag = false;
//        String[] table  = {"Array:0","EVEN:0","TPOT:0","EAC:0","Module:0"};
//        String[] new_table = {"Array:1","EVEN:0","TPOT:0","EAC:0","Module:0"};
//        for(int i = 0; i < ganttTaskMapper.getCapacity().size();i++){
//            use_power.put(ganttTaskMapper.getCapacity().get(i).getFactory_type(),String.valueOf(ganttTaskMapper.getCapacity().get(i).getProduct_in_ability())+","+String.valueOf(ganttTaskMapper.getCapacity().get(i).getProduct_out_ability()));
//        }
//        for(int i = 0; i < list_gantt.size();i++){
//            if(!Run_sheet.containsKey(list_gantt.get(i).getTarget())){
//                Run_sheet.put(list_gantt.get(i).getTarget(),table);
//                if(list_gantt.get(i).getFactory_type().equals("ARRAY"))
//                    Run_sheet.put(list_gantt.get(i).getTarget(), new String[]{"Array:1", "EVEN:0", "TPOT:0", "EAC:0", "Module:0"});
//                else if(list_gantt.get(i).getFactory_type().equals("EVEN"))
//                    Run_sheet.put(list_gantt.get(i).getTarget(), new String[]{"Array:0", "EVEN:1", "TPOT:0", "EAC:0", "Module:0"});
//                else if(list_gantt.get(i).getFactory_type().equals("TPOT"))
//                    Run_sheet.put(list_gantt.get(i).getTarget(), new String[]{"Array:0", "EVEN:0", "TPOT:1", "EAC:0", "Module:0"});
//                else if(list_gantt.get(i).getFactory_type().equals("EAC"))
//                    Run_sheet.put(list_gantt.get(i).getTarget(), new String[]{"Array:0", "EVEN:0", "TPOT:0", "EAC:1", "Module:0"});
//                else if(list_gantt.get(i).getFactory_type().equals("MODULE"))
//                    Run_sheet.put(list_gantt.get(i).getTarget(), new String[]{"Array:0", "EVEN:0", "TPOT:0", "EAC:0", "Module:1"});
////                    System.out.println(runGanttInfo.getTarget());
//            }
//            else {
//                new_table =  Run_sheet.get(list_gantt.get(i).getTarget());
//                if(list_gantt.get(i).getFactory_type().equals("ARRAY"))
//                    new_table[0] = "Array:1";
//                else if(list_gantt.get(i).getFactory_type().equals("EVEN"))
//                    new_table[1] = "EVEN:1";
//                else if(list_gantt.get(i).getFactory_type().equals("TPOT"))
//                    new_table[2] = "TPOT:1";
//                else if(list_gantt.get(i).getFactory_type().equals("EAC"))
//                    new_table[3] = "EAC:1";
//                else if(list_gantt.get(i).getFactory_type().equals("MODULE"))
//                    new_table[4] = "Module:1";
//                Run_sheet.put(list_gantt.get(i).getTarget(),new_table);
//            }
//        }
////        for (Map.Entry<String, String[]> entry : Run_sheet.entrySet()) {
////            System.out.println("Key = " + entry.getKey());
////            for(int i = 0 ; i < 5 ; i++){
////                System.out.println("Value = "+ entry.getValue()[i]);
////            }
////        }
//
//        System.out.println("!!!!!!!!!!!!!!!!!!!!!!!");
//        for(int i = 0; i < list_gantt.size();i++){
//            if(list_gantt.get(i).getFactory_type().equals("ARRAY")){
//                list_gantt_array.add(list_gantt.get(i));
//                System.out.println(list_gantt.get(i).getInput_time()+":"+list_gantt.get(i).getInput_amount());
//                System.out.println(list_gantt.get(i).getDuration());
//            }
//        }
//
//
//        for(int i = 0; i < list_gantt_array.size();i++){
//            if(map_time_array.containsKey(list_gantt_array.get(i).getInput_time())) {
//                map_time_array.put(list_gantt_array.get(i).getInput_time(), map_time_array.get(list_gantt_array.get(i).getInput_time()) + list_gantt_array.get(i).getInput_amount());
//            }
//            else
//                map_time_array.put(list_gantt_array.get(i).getInput_time(),list_gantt_array.get(i).getInput_amount());
//        }
//
////        for (Map.Entry<String, Double> entry : map_time_array.entrySet()) {
////            if(entry.getValue()>Double.parseDouble(use_power.get("ARRAY").split(",")[0])){
////                over_power_date.put(entry.getKey(),entry.getValue());
////            }
////        }
//        int count = 0;
//        System.out.println("----1111");
////        for (Map.Entry<String, Double> entry : map_time_array.entrySet()) {
////            System.out.println(entry.getKey()+":"+String.valueOf(entry.getValue()));
////        }
//        while (flag == false && count < 20){
//        count++;
//        flag = vaildCapacity(list_gantt,"ARRAY");
//        System.out.println(flag);
//        HashMap<String,Double> map_time_tmp = (HashMap<String,Double>)map_time_array.clone();
//            System.out.println("----------------------");
//            for (Map.Entry<String, Double> entry : map_time_array.entrySet()) {
//                System.out.println(entry.getKey()+":"+String.valueOf(entry.getValue()));
//            }
//        for (Map.Entry<String, Double> entry : map_time_tmp.entrySet()) {
//            if (flag) {
//                break;
//            }
//            if (entry.getValue() > Double.parseDouble(use_power.get("ARRAY").split(",")[0])) {
//                System.out.println("sssss");
////                System.out.println(entry.getKey() +":"+entry.getValue());
//                List<RunGanttInfo> tmp_target = new LinkedList<>();
//                for (int i = 0; i < list_gantt_array.size(); i++) {
//                    if (list_gantt_array.get(i).getInput_time().equals(entry.getKey())) {
//                        tmp_target.add(list_gantt_array.get(i));
////                        System.out.println(list_gantt_array.get(i));
//                    }
//                }
//                Collections.sort(tmp_target);
//                for (int i = 0; i < tmp_target.size(); i++) {
//                    String even_flag = Run_sheet.get(tmp_target.get(i).getTarget())[1].split(":")[1];
//                    String mod_flag = Run_sheet.get(tmp_target.get(i).getTarget())[4].split(":")[1];
//                    if (flag) {
//                        break;
//                    }
//                    if (even_flag.equals("1") || mod_flag.equals("1")) {
//                        String tmp_time = tmp_target.get(i).getInput_time();
//                        while(map_time_array.get(tmp_time) > Double.parseDouble(use_power.get("ARRAY").split(",")[0])){
//                            list_gantt = ganttTimeUpdate(list_gantt, tmp_target.get(i).getTarget(), tmp_target.get(i).getInput_time(), -1);
//                            if(!map_time_array.containsKey(tmp_target.get(i).getInput_time())){
//                                map_time_array.put(tmp_target.get(i).getInput_time(),0.0);
//                            }
//                            if (map_time_array.get(tmp_time) - Double.parseDouble(use_power.get("ARRAY").split(",")[0]) > 0) {
//                                if(map_time_array.get(tmp_target.get(i).getInput_time()) == 0.0){
//                                    System.out.println("9999999999999999999999");
//                                    System.out.println(tmp_target.get(i).getInput_time() + ":" + String.valueOf(map_time_array.get(tmp_target.get(i).getInput_time()) + Double.parseDouble(use_power.get("ARRAY").split(",")[0])));
//                                    map_time_array.put(tmp_target.get(i).getInput_time(), map_time_array.get(tmp_target.get(i).getInput_time()) + Double.parseDouble(use_power.get("ARRAY").split(",")[0]));
//                                    System.out.println(tmp_time + ":" + String.valueOf(map_time_array.get(tmp_time) - Double.parseDouble(use_power.get("ARRAY").split(",")[0])));
//                                    map_time_array.put(tmp_time, map_time_array.get(tmp_time) - Double.parseDouble(use_power.get("ARRAY").split(",")[0]));
//                                    tmp_target.get(i).setInput_amount(map_time_array.get(tmp_time));
//                                    Double tmp = calTimewithOut(list_gantt_array,tmp_target.get(i).getTarget(),map_time_array.get(tmp_time),tmp_target.get(i).getInput_time(),tmp_target.get(i).getDuration());
//                                    tmp_target.get(i).setInput_amount(tmp);
//                                    flag = vaildCapacity(list_gantt, "ARRAY");
//                                    if(flag)
//                                        break;
//                                }
//                                else {
//                                    System.out.println("8888888888888888888888888");
//                                    Double cap = Double.parseDouble(use_power.get("ARRAY").split(",")[0]) - (map_time_array.get(tmp_target.get(i).getInput_time())%Double.parseDouble(use_power.get("ARRAY").split(",")[0]));
//                                    System.out.println(cap);
//                                    map_time_array.put(tmp_target.get(i).getInput_time(), map_time_array.get(tmp_target.get(i).getInput_time()) + cap);
//                                    map_time_array.put(tmp_time, map_time_array.get(tmp_time) - cap);
////                                    tmp_target.get(i).setInput_amount(map_time_array.get(tmp_time));
//                                    Double tmp = calTimewithOut(list_gantt_array,tmp_target.get(i).getTarget(),map_time_array.get(tmp_time),tmp_target.get(i).getInput_time(),tmp_target.get(i).getDuration());
//                                    tmp_target.get(i).setInput_amount(tmp);
//                                    flag = vaildCapacity(list_gantt, "ARRAY");
//                                    if(flag)
//                                        break;
//                                }
//                            }
//                        }
////                        if(map_time_array.get(tmp_time) - Double.parseDouble(use_power.get("ARRAY").split(",")[0]) > 0){
////                            System.out.println("!!!!!!!!!!");
////                            System.out.println(tmp_time);
////                        }
//                        if(map_time_array.get(tmp_time) <= Double.parseDouble(use_power.get("ARRAY").split(",")[0]) && map_time_array.get(tmp_target.get(i).getInput_time()) <= Double.parseDouble(use_power.get("ARRAY").split(",")[0])){
//                            break;
//                        }
//                    }
////                    else {
//////                    System.out.println(tmp_target.get(i).getTarget() + " " + tmp_target.get(i).getInput_time() + "+1");
////                        String tmp_time = tmp_target.get(i).getInput_time();
////                        list_gantt = ganttTimeUpdate(list_gantt, tmp_target.get(i).getTarget(), tmp_target.get(i).getInput_time(), 1);
//////                        calendar.setTime(format.parse(tmp_target.get(i).getInput_time()));
//////                        calendar.add(calendar.DATE, 1);
////                        if(!map_time_array.containsKey(tmp_target.get(i).getInput_time())){
////                            map_time_array.put(tmp_target.get(i).getInput_time(),0.0);
////                        }
////                        Double tmp = map_time_array.get(tmp_time) - Double.parseDouble(use_power.get("ARRAY").split(",")[0]);
////                        map_time_array.put(tmp_target.get(i).getInput_time(), map_time_array.get(tmp_target.get(i).getInput_time()) + tmp);
////                        map_time_array.put(tmp_time, Double.parseDouble(use_power.get("ARRAY").split(",")[0]));
////                        flag = vaildCapacity(list_gantt, "ARRAY");
////                        System.out.println(flag);
////                    }
//                }
//
//                if(flag==true){
//                    System.out.println("----0000");
//                    for (Map.Entry<String, Double> entry2 : map_time_array.entrySet()) {
//                        System.out.println(entry2.getKey()+":"+String.valueOf(entry2.getValue()));
//                    }
//                }
//
//            }
//        }
//        }
//    }
    public List<RunGanttInfo>ganttTimeUpdate(List<RunGanttInfo> list_gannt, String Target,String Time,int flag) throws ParseException {
        DateFormat format=new SimpleDateFormat("yyyy/MM/dd");
        Calendar calendar = Calendar.getInstance();
        for(int i = 0; i < list_gannt.size();i ++){
            if (list_gannt.get(i).getTarget().equals(Target) && list_gannt.get(i).getInput_time().equals(Time)){
                calendar.setTime(format.parse(Time));
                calendar.add(calendar.DATE,flag * 1);
                list_gannt.get(i).setInput_time(format.format(calendar.getTime()));
                list_gannt.get(i).setDuration(list_gannt.get(i).getDuration()+1);
            }
        }
        return list_gannt;
    }

    public Boolean vaildCapacity(List<RunGanttInfo> list_gantt, String Factory) throws ParseException {
        HashMap<String,String> use_power = new HashMap<>();
        HashMap<String,Double> map_time = new HashMap<>();
        DateFormat format=new SimpleDateFormat("yyyy/MM/dd");
        Calendar calendar = Calendar.getInstance();
        Boolean flag = true;
        List<GanttCapacity> ganttCapacityList = ganttTaskMapper.getCapacity();
        for(int i = 0; i < ganttCapacityList.size();i++){
            use_power.put(ganttCapacityList.get(i).getFactory_type(),String.valueOf(ganttCapacityList.get(i).getProduct_in_ability())+","+String.valueOf(ganttCapacityList.get(i).getProduct_out_ability()));
        }
        for(int i = 0; i < list_gantt.size();i++){
            if(list_gantt.get(i).getFactory_type().equals(Factory)){
                System.out.println(list_gantt.get(i).getInput_time()+":"+list_gantt.get(i).getInput_amount());
                System.out.println(list_gantt.get(i).getDuration());
                if(map_time.containsKey(list_gantt.get(i).getInput_time())) {
                    if(list_gantt.get(i).getDuration() == 0){
                        map_time.put(list_gantt.get(i).getInput_time(), map_time.get(list_gantt.get(i).getInput_time()) + list_gantt.get(i).getInput_amount());
                        if(map_time.get(list_gantt.get(i).getInput_time())>Double.parseDouble(use_power.get(Factory).split(",")[0])){
//                            System.out.println("3333333333333333333333333333333333333333");
//                            System.out.println(list_gantt.get(i).getInput_time());
//                            System.out.println(map_time.get(list_gantt.get(i).getInput_time()));
    //                        return false;
                                flag = false;
                            }
                    }
                    else{
//                        System.out.println("4444444444444444");
                        if(list_gantt.get(i).getDuration() > 0){
                            System.out.println(list_gantt.get(i).getInput_time());
                        }
//                        System.out.println("555555555555555555");
                        calendar.setTime(format.parse(list_gantt.get(i).getInput_time()));
                        if(Factory.equals("ARRAY")){
                            calendar.add(calendar.DATE,1*list_gantt.get(i).getDuration());
                        }
                        else
                            calendar.add(calendar.DATE,-1*list_gantt.get(i).getDuration());
                        map_time.put(format.format(calendar.getTime()), map_time.get(list_gantt.get(i).getInput_time()) + list_gantt.get(i).getInput_amount());

                    }

                }
                else{
                    map_time.put(list_gantt.get(i).getInput_time(),list_gantt.get(i).getInput_amount());
                    if(map_time.get(list_gantt.get(i).getInput_time())>Double.parseDouble(use_power.get(Factory).split(",")[0])){
//                        System.out.println("2222222222222222222222222222");
//                        System.out.println(list_gantt.get(i).getInput_time());
//                        System.out.println(map_time.get(list_gantt.get(i).getInput_time()));
//                        return false;
                        flag = false;
                    }
                }

            }
        }
        return flag;
    }

    public Double calTimewithOut(List<RunGanttInfo> list_gantt,String target,Double sum ,String time, int duration) throws ParseException {
        DateFormat format=new SimpleDateFormat("yyyy/MM/dd");
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(format.parse(time));
        calendar.add(calendar.DATE,duration* 1);
        String new_time = format.format(calendar.getTime());
        for(int i = 0; i < list_gantt.size();i++){
            if(list_gantt.get(i).getInput_time().equals(new_time) && !list_gantt.get(i).getTarget().equals(target)){
                sum -= list_gantt.get(i).getInput_amount();
            }
        }
        return sum;
    }

//    public Result setMacroGantt(String fileName,int pilot) throws Exception {
      public Result setMacroGantt(MacroGantt_Status macroGantt_status_new) throws Exception {
        ganttTaskMapper.deleteMacroStatus();
        ganttTaskMapper.insertMacroStatus(macroGantt_status_new);
        MacroGantt_Status macroGantt_status = ganttTaskMapper.getMacroStatus();
//          String depart = "实验";
//          String project = "Pilot（整合实验）";
//          String pilot = "Pilot 4";
//          String target = "OLED器件验证";
        String depart = macroGantt_status.getDepart();
        String project = macroGantt_status.getProject();
        String pilot = macroGantt_status.getPilot();
        String target = macroGantt_status.getTarget();
        String fileName = "D:\\yunying\\upload\\excel\\gantt.xlsx";
        String sheetName = "编码规则";
          Excel_Util.workbook = new XSSFWorkbook(fileName);
          MacroGantt macroGantt = new MacroGantt();
          MacroGantt fathertMacro_in = new MacroGantt();
          MacroGantt fathertMacro_out = new MacroGantt();
          DateFormat format=new SimpleDateFormat("yyyy-MM-dd");
          String cur_sheet_name = "";
          HashMap<Integer,String>facotry_name = new HashMap<>();
        ganttTaskMapper.deleteMacroTask();
        int rela_row = Excel_Util.workbook.getSheet("编码规则").getLastRowNum();
        HashMap<String, String> code_Fir = new HashMap<>();
        HashMap<String, String> code_Sec = new HashMap<>();
        HashMap<String, String> code_Thi = new HashMap<>();
        HashMap<String, String> code_Fou = new HashMap<>();
        String tmp_str = "*";
        int tmp_int = 1;
        //将编码规则读入字典
        while (!tmp_str.equals("") && tmp_int <= rela_row) {
            tmp_str = Excel_Util.readExcelData(fileName, sheetName, tmp_int, 0);
            System.out.println(tmp_str);
            System.out.println(Excel_Util.readExcelData(fileName, sheetName, tmp_int, 1));
            code_Fir.put(Excel_Util.readExcelData(fileName, sheetName, tmp_int, 1), Excel_Util.readExcelData(fileName, sheetName, tmp_int, 0));
            tmp_int++;
        }
        for(Map.Entry<String, String> entry : code_Fir.entrySet()){
            if(entry.getValue().equals(depart)){
                depart = entry.getKey();
            }
        }
        tmp_int = 1;
        tmp_str = "*";
        while (!tmp_str.equals("") && tmp_int <= rela_row) {
            tmp_str = Excel_Util.readExcelData(fileName, sheetName, tmp_int, 2);
            System.out.println(tmp_str);
            System.out.println(Excel_Util.readExcelData(fileName, sheetName, tmp_int, 3));
            code_Sec.put(Excel_Util.readExcelData(fileName, sheetName, tmp_int, 3), Excel_Util.readExcelData(fileName, sheetName, tmp_int, 2));
            tmp_int++;
        }
          for(Map.Entry<String, String> entry : code_Sec.entrySet()){
              if(entry.getValue().equals(project)){
                  project = entry.getKey();
              }
          }
        tmp_int = 1;
        tmp_str = "*";

        while (!tmp_str.equals("") && tmp_int <= rela_row) {
            tmp_str = Excel_Util.readExcelData(fileName, sheetName, tmp_int, 4);
            System.out.println(tmp_str);
            System.out.println(Excel_Util.readExcelData(fileName, sheetName, tmp_int, 5));
            code_Thi.put(Excel_Util.readExcelData(fileName, sheetName, tmp_int, 5), Excel_Util.readExcelData(fileName, sheetName, tmp_int, 4));
            tmp_int++;
        }
          for(Map.Entry<String, String> entry : code_Thi.entrySet()){
              if(entry.getValue().equals(pilot)){
                  pilot = entry.getKey();
              }
          }
        tmp_int = 1;
        tmp_str = "*";

        while (!tmp_str.equals("") && tmp_int <= rela_row) {
            tmp_str = Excel_Util.readExcelData(fileName, sheetName, tmp_int, 6);
            System.out.println(tmp_str);
            System.out.println(Excel_Util.readExcelData(fileName, sheetName, tmp_int, 7));
            code_Fou.put(Excel_Util.readExcelData(fileName, sheetName, tmp_int, 7), Excel_Util.readExcelData(fileName, sheetName, tmp_int, 6));
            tmp_int++;
        }
          for(Map.Entry<String, String> entry : code_Fou.entrySet()){
              if(entry.getValue().equals(target)){
                  target = entry.getKey();
              }
          }
          String[] tmp = {depart,project,pilot,target};
          for(int i = 0; i <tmp.length;i++){
              System.out.println("-----------");
              System.out.println(tmp[i]);
          }
          System.out.println(tmp);
        facotry_name.put(1,"Array");
        facotry_name.put(2,"EVEN");
        facotry_name.put(3,"TPOT");
        facotry_name.put(4,"EAC");
        facotry_name.put(5,"MODULE");
        for(int i = 1; i <= 5; i++){
            fathertMacro_in.setId(facotry_name.get(i)+"投入");
            fathertMacro_in.setText(facotry_name.get(i)+"投入");
            fathertMacro_in.setGantt_type(facotry_name.get(i));
            fathertMacro_in.setColor("rgba(0,0,0,0)");
            fathertMacro_in.setRender("split");
            fathertMacro_in.setNumber(i*2);
            fathertMacro_out.setId(facotry_name.get(i)+"产出");
            fathertMacro_out.setText(facotry_name.get(i)+"产出");
            fathertMacro_out.setGantt_type(facotry_name.get(i));
            fathertMacro_out.setColor("rgba(0,0,0,0)");
            fathertMacro_out.setRender("split");
            fathertMacro_out.setNumber(i*2+1);
            ganttTaskMapper.insertMacroGantt(fathertMacro_in);
            ganttTaskMapper.insertMacroGantt(fathertMacro_out);
        }
        HashMap<Integer, String> sheet_name = new HashMap<>();
        sheet_name.put(1, "ARRAY计划");
        sheet_name.put(2, "EVEN计划");
        sheet_name.put(3, "TPOT计划");
        sheet_name.put(4, "EAC计划");
        sheet_name.put(5, "MODULE计划");
        for(int sheet_index = 1; sheet_index <= 5; sheet_index++){
            cur_sheet_name = sheet_name.get(sheet_index);
            int cur_row = Excel_Util.readrowNum(fileName, cur_sheet_name);
            int cur_col = Excel_Util.readcolNum(fileName, cur_sheet_name);
            int cur_type = Excel_Util.readWantCol(fileName, cur_sheet_name, 0, "IN/OUT");
            int cur_mark = Excel_Util.readWantCol(fileName, cur_sheet_name, 0, "MARK");
            for(int col = cur_type+1; col< cur_col;col++){
                String cur_date = Excel_Util.readExcelData(fileName,cur_sheet_name,0,col);
                if(cur_date.isEmpty())
                    break;
                Double cur_value_in = 0.0;
                Double cur_value_out = 0.0;
                for(int row = 1; row < cur_row; row++){
                    String tmp_date = Excel_Util.readExcelData(fileName,cur_sheet_name,row,col);
                    String mark = Excel_Util.readExcelData(fileName, cur_sheet_name, row, cur_mark);
                    Boolean flag = true;
                    if(!tmp_date.equals("")&&!tmp_date.equals("0")){
//                        if(pilot == "-1")
//                            cur_value_in += Double.parseDouble(tmp_date);
//                        else{
//                            if(pilot == String.valueOf(mark.charAt(2))){
//                                cur_value_in += Double.parseDouble(tmp_date);
//                            }
//                            else{
//                                row++;
//                                continue;
//                            }
//                        }
                        if(tmp[0].equals("-1")&&tmp[1].equals("-1")&&tmp[2].equals("-1")&&tmp[3].equals("-1")){
                            cur_value_in += Double.parseDouble(tmp_date);
                        }
                        else{
                            for(int tmp_i = 0; tmp_i < 4; tmp_i++){
                                if(!tmp[tmp_i].equals("-1")){
                                    if(tmp[tmp_i].equals(String.valueOf(mark.charAt(tmp_i)))){
                                        continue;
                                    }
                                    else{
                                        flag = false;
                                        break;
                                    }

                                }
                            }
                            if(flag == true)
                                cur_value_in += Double.parseDouble(tmp_date);
                            else {
                                row++;
                                continue;
                            }
                        }
                    }
                    row++;
                }
                for(int row = 2; row < cur_row; row++){
                    String tmp_date = Excel_Util.readExcelData(fileName,cur_sheet_name,row,col);
                    String mark = Excel_Util.readExcelData(fileName, cur_sheet_name, row, cur_mark);
                    Boolean flag = true;
                    if(!tmp_date.equals("")&&!tmp_date.equals("0")){
//                        if(pilot == "-1")
//                            cur_value_out += Double.parseDouble(tmp_date);
//                        else{
//                            if(pilot == String.valueOf(mark.charAt(2))){
//                                cur_value_out += Double.parseDouble(tmp_date);
//                            }
//                            else{
//                                row++;
//                                continue;
//                            }
//                        }
                        if(tmp[0].equals("-1")&&tmp[1].equals("-1")&&tmp[2].equals("-1")&&tmp[3].equals("-1")){
                            cur_value_out += Double.parseDouble(tmp_date);
                        }
                        else{
                            for(int tmp_i = 0; tmp_i < 4; tmp_i++){
                                if(!tmp[tmp_i].equals("-1")){
                                    if(tmp[tmp_i].equals(String.valueOf(mark.charAt(tmp_i)))){
                                        continue;
                                    }
                                    else{
                                        flag = false;
                                        break;
                                    }

                                }
                            }
                            if(flag == true)
                                cur_value_out += Double.parseDouble(tmp_date);
                            else {
                                row++;
                                continue;
                            }
                        }

                    }
                    row++;
                }
                if(cur_value_in > 0.0){
                    macroGantt.setDuration(1);
                    macroGantt.setParent(facotry_name.get(sheet_index)+"投入");
                    macroGantt.setUse_amount(String.valueOf(cur_value_in));
                    macroGantt.setColor("rgba(255,165,0,0.5)");
                    macroGantt.setStart_date(Excel_Util.DateToFormat(cur_date));
                    macroGantt.setId(sheet_name.get(sheet_index)+String.valueOf(col)+"_in");
                    ganttTaskMapper.insertMacroGantt(macroGantt);
                }
                if(cur_value_out > 0.0){
                    macroGantt.setDuration(1);
                    macroGantt.setParent(facotry_name.get(sheet_index)+"产出");
                    macroGantt.setUse_amount(String.valueOf(cur_value_out));
                    macroGantt.setColor("rgba(192,192,192,0.5)");
                    macroGantt.setStart_date(Excel_Util.DateToFormat(cur_date));
                    macroGantt.setId(sheet_name.get(sheet_index)+String.valueOf(col)+"_out");
                    ganttTaskMapper.insertMacroGantt(macroGantt);
                }

                System.out.println(sheet_name.get(sheet_index)+"投入:"+Excel_Util.DateToFormat(cur_date)+":"+cur_value_in);
                System.out.println(sheet_name.get(sheet_index)+"产出:"+Excel_Util.DateToFormat(cur_date)+":"+cur_value_out);

//                    System.out.println(Excel_Util.DateToFormat(Excel_Util.readExcelData(fileName,cur_sheet_name,0,col)));
            }
        }


        return new Result("200","success");
    }


}
