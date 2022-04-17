package com.tianma.yunying.test;

import com.tianma.yunying.entity.GanttCapacity;
import com.tianma.yunying.entity.Result;
import com.tianma.yunying.entity.RunGanttInfo;
import com.tianma.yunying.entity.RunGanttTask;
import com.tianma.yunying.mapper.GanttTaskMapper;
import com.tianma.yunying.util.Excel_Util;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;

import javax.swing.filechooser.FileSystemView;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.NumberFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;


public class test324 {
    @Autowired
    public static GanttTaskMapper ganttTaskMapper;
    public static void main(String[] args) throws Exception {
        String filePath = "D:\\CodePath\\test\\new_data.xlsx";
        Excel_Util.workbook = new XSSFWorkbook(filePath);
        List<RunGanttInfo> list_gantt = new LinkedList<>();
        HashMap<Integer,String> sheet_name = new HashMap<>();
        HashMap<String,String[]> Run_sheet = new HashMap<>();
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
            int col_output_amount = Excel_Util.readWantCol(filePath,sheetName,0,"产出数量");
            int col_number = Excel_Util.readWantCol(filePath,sheetName,0,"优先级");
            int col_prodect = Excel_Util.readWantCol(filePath,sheetName,0,"产品型号");
            int col_desc = Excel_Util.readWantCol(filePath,sheetName,0,"说明");
            int col_cycle = Excel_Util.readWantCol(filePath,sheetName,0,"cycle");
            int col_bank = Excel_Util.readWantCol(filePath,sheetName,0,"bank");
            int cur_rownum = Excel_Util.readrowNum(filePath,sheetName);
            for(int row = 1; row <= cur_rownum; row++){
                RunGanttInfo runGanttInfo = new RunGanttInfo();
                runGanttInfo.setTarget(Excel_Util.readExcelData(filePath,sheetName,row,col_target));
                runGanttInfo.setDepartment(Excel_Util.readExcelData(filePath,sheetName,row,col_department));
                runGanttInfo.setCustomer(Excel_Util.readExcelData(filePath,sheetName,row,col_customer));
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
                if(!Excel_Util.readExcelData(filePath,sheetName,row,col_output_amount).equals("")){
                    runGanttInfo.setOutput_amount(Double.parseDouble(Excel_Util.readExcelData(filePath,sheetName,row,col_output_amount)));
                    NumberFormat num = NumberFormat.getPercentInstance();
                    String rate_yield = num.format(runGanttInfo.getOutput_amount()/runGanttInfo.getInput_amount());
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


        list_gantt = infoUpdate(list_gantt,Run_sheet,readInputEven(filePath));
        for(int i = 0; i < list_gantt.size();i++){
//            ganttTaskMapper.insertRunGanttInfo(list_gantt.get(i));
            System.out.println(list_gantt.get(i));
        }
//
        toGantt(list_gantt);
    }
    public static HashMap<String,String[]> readInputEven(String fileName) throws Exception {
        String filePath = fileName;
        HashMap<String,String[]> EVEN_Sheet = new HashMap<>();
//        String filePath = "D:\\CodePath\\test\\副本2022年M+3实验需求Rev 03-整0318.xlsx";
        Excel_Util.workbook = new XSSFWorkbook(filePath);
        DateFormat format=new SimpleDateFormat("yyyy/MM/dd");
        int col_target = Excel_Util.readWantCol(filePath,"EVEN输入表",0,"实验目的");
        int col_input = Excel_Util.readWantCol(filePath,"EVEN输入表",0,"投入时间");
        int col_output = col_input+1;
        int col_in_amount = col_input+2;
        int col_out_amount = col_input+3;

        for(int row = 1; row <= Excel_Util.readrowNum(filePath,"EVEN输入表");row++){
            String input = Excel_Util.DateToFormat(Excel_Util.readExcelData(filePath,"EVEN输入表",row,col_input));
            String output = Excel_Util.DateToFormat(Excel_Util.readExcelData(filePath,"EVEN输入表",row,col_output));
            String input_amount = Excel_Util.readExcelData(filePath,"EVEN输入表",row,col_in_amount);
            String output_amount = Excel_Util.readExcelData(filePath,"EVEN输入表",row,col_out_amount);
            EVEN_Sheet.put(Excel_Util.readExcelData(filePath,"EVEN输入表",row,col_target), new String[]{input, output,input_amount,output_amount});
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

    public static HashMap<String,String[]> readInputEven() throws Exception {
        HashMap<String,String[]> EVEN_Sheet = new HashMap<>();
        String filePath = "D:\\CodePath\\test\\副本2022年M+3实验需求Rev 03-整0318.xlsx";
        Excel_Util.workbook = new XSSFWorkbook(filePath);
        DateFormat format=new SimpleDateFormat("yyyy/MM/dd");
        int col_target = Excel_Util.readWantCol(filePath,"EVEN输入表",0,"实验目的");
        int col_input = Excel_Util.readWantCol(filePath,"EVEN输入表",0,"投入时间");
        int col_output = col_input+1;
        for(int row = 1; row <= Excel_Util.readrowNum(filePath,"EVEN输入表");row++){
            String input = Excel_Util.DateToFormat(Excel_Util.readExcelData(filePath,"EVEN输入表",row,col_input));
            String output = Excel_Util.DateToFormat(Excel_Util.readExcelData(filePath,"EVEN输入表",row,col_output));
            EVEN_Sheet.put(Excel_Util.readExcelData(filePath,"EVEN输入表",row,col_target), new String[]{input, output});
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

    public static List<RunGanttInfo> infoUpdate(List<RunGanttInfo> list_gantt,HashMap<String,String[]> Run_sheet,HashMap<String,String[]> EVEN_Sheet) throws ParseException {
        HashMap<String,String> factory_count = new HashMap<>();
        Calendar rightNow = Calendar.getInstance();
        DateFormat format=new SimpleDateFormat("yyyy/MM/dd");
        factory_count.put("Array","1");
        factory_count.put("EVEN","1");
        factory_count.put("TPOT","1");
        factory_count.put("EAC","1");
        factory_count.put("Module","1");
        String[] use_power = {"300,25,2","300,3,1","300,3,1","300,3,1","28000,5,1"};
        for (Map.Entry<String, String[]> entry : Run_sheet.entrySet()) {
            System.out.println("Key = " + entry.getKey());
            String even_flag = entry.getValue()[1].split(":")[1];
            String mod_flag = entry.getValue()[4].split(":")[1];
            for(int i = 0; i <list_gantt.size();i++){
                if(list_gantt.get(i).getTarget().equals(entry.getKey())&&!entry.getKey().equals("")){
                    if(even_flag.equals("1")){
                        String even_input = EVEN_Sheet.get(entry.getKey())[0];
                        String even_output = EVEN_Sheet.get(entry.getKey())[1];
                        list_gantt.get(i).setInput_time(even_input);
                        list_gantt.get(i).setOutput_time(even_output);
                        if(list_gantt.get(i).getFactory_type().equals("ARRAY")){
//                            System.out.println(list_gantt.get(i));
                            String inputTime = EVEN_Sheet.get(entry.getKey())[0];
                            String outputTime = EVEN_Sheet.get(entry.getKey())[1];
                            String cycleTime = use_power[0].split(",")[1];
                            String bankTime  = use_power[0].split(",")[2];
                            Date date = format.parse(inputTime);
                            rightNow.setTime(date);
                            rightNow.add(Calendar.DAY_OF_YEAR, -1*Integer.parseInt(bankTime));
                            String update_output = format.format(rightNow.getTime());
                            rightNow.setTime(rightNow.getTime());
                            rightNow.add(Calendar.DAY_OF_YEAR, -1*Integer.parseInt(cycleTime));
                            String update_input = format.format(rightNow.getTime());
                            list_gantt.get(i).setOutput_time(update_output);
                            list_gantt.get(i).setInput_time(update_input);
//                            System.out.println(list_gantt.get(i));
                        }
                        else if(list_gantt.get(i).getFactory_type().equals("TPOT")){
//                            System.out.println(list_gantt.get(i));
                            String inputTime = EVEN_Sheet.get(entry.getKey())[0];
                            String outputTime = EVEN_Sheet.get(entry.getKey())[1];
                            String cycleTime = use_power[2].split(",")[1];
                            String bankTime  = use_power[1].split(",")[2];
                            Date date = format.parse(outputTime);
                            rightNow.setTime(date);
                            rightNow.add(Calendar.DAY_OF_YEAR, Integer.parseInt(bankTime));
                            String update_input = format.format(rightNow.getTime());
                            rightNow.setTime(rightNow.getTime());
                            rightNow.add(Calendar.DAY_OF_YEAR, Integer.parseInt(cycleTime));
                            String update_output = format.format(rightNow.getTime());
                            list_gantt.get(i).setOutput_time(update_output);
                            list_gantt.get(i).setInput_time(update_input);
//                            System.out.println(list_gantt.get(i));
                        }
                        else if(list_gantt.get(i).getFactory_type().equals("EAC")){
//                            System.out.println(list_gantt.get(i));
                            String inputTime = EVEN_Sheet.get(entry.getKey())[0];
                            String outputTime = EVEN_Sheet.get(entry.getKey())[1];
                            String cycleTime_ForTpot = use_power[2].split(",")[1];
                            String bankTime_ForTpot  = use_power[1].split(",")[2];
                            String cycleTime_ForEAC = use_power[3].split(",")[1];
                            String bankTime_ForEAC  = use_power[2].split(",")[2];
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
                            String cycleTime_ForTpot = use_power[2].split(",")[1];
                            String bankTime_ForTpot  = use_power[1].split(",")[2];
                            String cycleTime_ForEAC = use_power[3].split(",")[1];
                            String bankTime_ForEAC  = use_power[2].split(",")[2];
                            String cycleTime_ForModule = use_power[4].split(",")[1];
                            String bankTime_ForModule  = use_power[3].split(",")[2];
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
//                                System.out.println(list_gantt.get(i));
                                String cycleTime = use_power[3].split(",")[1];
                                String bankTime  = use_power[3].split(",")[2];
                                Date date = format.parse(inputTime);
                                rightNow.setTime(date);
                                rightNow.add(Calendar.DAY_OF_YEAR, -1*Integer.parseInt(bankTime));
                                String update_output = format.format(rightNow.getTime());
                                rightNow.setTime(rightNow.getTime());
                                rightNow.add(Calendar.DAY_OF_YEAR, -1*Integer.parseInt(cycleTime));
                                String update_input = format.format(rightNow.getTime());
                                list_gantt.get(i).setOutput_time(update_output);
                                list_gantt.get(i).setInput_time(update_input);
                                System.out.println(list_gantt.get(i));
                            }

                            if(list_gantt.get(i).getFactory_type().equals("TPOT")){
//                                System.out.println(list_gantt.get(i));
                                String cycleTime_EAC = use_power[3].split(",")[1];
                                String bankTime_EAC  = use_power[3].split(",")[2];
                                String cycleTime_Tpot = use_power[2].split(",")[1];
                                String bankTime_Tpot  = use_power[2].split(",")[2];
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
                                String cycleTime_EAC = use_power[3].split(",")[1];
                                String bankTime_EAC  = use_power[3].split(",")[2];
                                String cycleTime_Tpot = use_power[2].split(",")[1];
                                String bankTime_Tpot  = use_power[2].split(",")[2];
                                String cycleTime_EVEN = use_power[1].split(",")[1];
                                String bankTime_EVEN  = use_power[1].split(",")[2];
                                String cycleTime_Array = use_power[0].split(",")[1];
                                String bankTime_Array  = use_power[0].split(",")[2];
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
            }
//            System.out.println("----------------");
        }
        return list_gantt;
    }

    public static void toGantt(List<RunGanttInfo> list_gantt){
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
//                ganttTaskMapper.insertTask_run(father);
                System.out.println(father);
            }
            parent.setColor("rgba(0,0,0,0)");
            parent.setParent(list_gantt.get(i).getTarget()+"father");
            parent.setText(list_gantt.get(i).getTarget());
            parent.setOpen("true");
            parent.setRender("spilt");
            parent.setId(list_gantt.get(i).getTarget());
            parent.setFactory_type(list_gantt.get(i).getFactory_type());
            System.out.println(parent);
//            ganttTaskMapper.insertTask_run(parent);
            System.out.println(parent);
            runGantttask_input.setParent(parent.getId());
            runGantttask_input.setId(parent.getId()+"投入");
            runGantttask_input.setColor("rgba(255,165,0,0.5)");
            runGantttask_input.setStart_date(list_gantt.get(i).getInput_time());
            runGantttask_input.setUse_amount(list_gantt.get(i).getInput_amount());
            System.out.println(runGantttask_input);
//            ganttTaskMapper.insertTask_run(runGantttask_input);
            System.out.println(runGantttask_input);
            runGantttask_output.setParent(parent.getId());
            runGantttask_output.setId(parent.getId()+"产出");
            runGantttask_output.setColor("rgba(192,192,192,0.5)");
            runGantttask_output.setStart_date(list_gantt.get(i).getOutput_time());
            runGantttask_output.setUse_amount(list_gantt.get(i).getOutput_amount());
            System.out.println(runGantttask_output);
//            ganttTaskMapper.insertTask_run(runGantttask_output);
            System.out.println(runGantttask_output);
        }

    }

    public static void writePlanExcel(List<RunGanttInfo> list_gantt) throws IOException, ParseException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet_array = workbook.createSheet("Array计划");
        Sheet sheet_even = workbook.createSheet("EVEN计划");
        Sheet sheet_tpot = workbook.createSheet("TPOT计划");
        Sheet sheet_eac = workbook.createSheet("EAC计划");
        Sheet sheet_module = workbook.createSheet("MODULE计划");
        DateFormat df = new SimpleDateFormat("yyyy/MM/dd");
        List<GanttCapacity> test_Capacity = new LinkedList<>();
        HashMap<Integer,Sheet> map_sheet = new HashMap();
        HashMap<Integer,String> factory_mark = new HashMap();
        HashMap<Integer,List<RunGanttInfo>> factory_gantt = new HashMap();
        HashMap<Integer,GanttCapacity> test_CapaCity = new HashMap<>();
        GanttCapacity array = new GanttCapacity();
        GanttCapacity even  = new GanttCapacity();
        GanttCapacity tpot = new GanttCapacity();
        GanttCapacity eac = new GanttCapacity();
        GanttCapacity module = new GanttCapacity();
        array.setFactory_type("Array");
        array.setProduct_in_ability(300.0);
        array.setProduct_out_ability(300.0);
        even.setFactory_type("EVEN");
        even.setProduct_in_ability(300.0);
        even.setProduct_out_ability(300.0);
        tpot.setFactory_type("TPOT");
        tpot.setProduct_in_ability(300.0);
        tpot.setProduct_out_ability(300.0);
        eac.setFactory_type("EVEN");
        eac.setProduct_in_ability(300.0);
        eac.setProduct_out_ability(28000.0);
        module.setFactory_type("MODULE");
        module.setProduct_in_ability(28000.0);
        module.setProduct_out_ability(28000.0);
        test_CapaCity.put(0,array);
        test_CapaCity.put(1,even);
        test_CapaCity.put(2,tpot);
        test_CapaCity.put(3,eac);
        test_CapaCity.put(4,module);

        List<RunGanttInfo> list_gantt_array = new LinkedList<>();
        List<RunGanttInfo> list_gantt_even = new LinkedList<>();
        List<RunGanttInfo> list_gantt_tpot = new LinkedList<>();
        List<RunGanttInfo> list_gantt_eac = new LinkedList<>();
        List<RunGanttInfo> list_gantt_module = new LinkedList<>();

        for(int i = 0; i <list_gantt.size();i++){
            if(list_gantt.get(i).getFactory_type().equals("Array"))
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
        factory_mark.put(0,"Array");
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
                        for(int j = 0; j <Math.ceil(cur_gantt.get(i).getInput_amount()/test_CapaCity.get(index).getProduct_in_ability());j++){
                            cell_in = row_in.createCell(cellnum);
                            if(tmp_amount - test_CapaCity.get(index).getProduct_in_ability() < 0){
                                cell_in = row_in.createCell(cellnum);
                                cell_in.setCellValue(tmp_amount);
                            }
                            else if(tmp_amount - test_CapaCity.get(index).getProduct_in_ability() >0){
                                System.out.println("87777777887887");
                                System.out.println(cur_gantt.get(i).getTarget());
                                tmp_amount -= test_CapaCity.get(index).getProduct_in_ability();
                                cell_in.setCellValue(test_CapaCity.get(index).getProduct_in_ability());
                                cellnum++;
                            }
                        }
                    }
                    if(df.format(cellValue).equals(cur_gantt.get(i).getOutput_time()))
                    {
                        Double tmp_amount = cur_gantt.get(i).getOutput_amount();
                        for(int j = 0; j <Math.ceil(cur_gantt.get(i).getOutput_amount()/test_CapaCity.get(index).getProduct_out_ability());j++){

                            cell_out = row_out.createCell(cellnum);
                            if(tmp_amount - test_CapaCity.get(index).getProduct_out_ability() < 0){
                                cell_out = row_out.createCell(cellnum);
                                cell_out.setCellValue(tmp_amount);
                            }
                            else if(tmp_amount - test_CapaCity.get(index).getProduct_out_ability() >0){
                                System.out.println("87777777887887产出");
                                System.out.println(cur_gantt.get(i).getTarget());
                                tmp_amount -= test_CapaCity.get(index).getProduct_out_ability();
                                cell_out.setCellValue(test_CapaCity.get(index).getProduct_out_ability());
                                cellnum++;
                            }
                        }
                    }
                }
            }
        }

        String fileNmae="运营部生成计划" + ".xlsx";
        File desktopDir= FileSystemView.getFileSystemView().getHomeDirectory();//获取桌面的目录
        String desktopPath=desktopDir.getAbsolutePath();//获取桌面的绝对路径
        String filePath=desktopPath+"\\"+fileNmae;
        FileOutputStream out=new FileOutputStream(filePath);
        workbook.write(out);
    }
}
