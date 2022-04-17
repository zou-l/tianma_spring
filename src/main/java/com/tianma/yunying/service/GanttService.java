//package com.tianma.yunying.service;
//
//import com.tianma.yunying.entity.Gantt_Detail;
//import com.tianma.yunying.entity.Gantt_Info;
//import com.tianma.yunying.mapper.GanttMapper;
//import com.tianma.yunying.test.PoiExcelTest;
//import org.springframework.beans.factory.annotation.Autowired;
//import org.springframework.stereotype.Service;
//
//import java.text.DateFormat;
//import java.text.SimpleDateFormat;
//import java.util.Date;
//import java.util.HashMap;
//
//@Service
//public class GanttService {
//    @Autowired
//    GanttMapper ganttMapper;
//
//    public void getInfo() throws Exception {
//        String filelName = "D:\\CodePath\\test\\联排计划展示基础表 (1).xlsx";
//        HashMap<Integer, String> sheet_name = new HashMap<>();
//        sheet_name.put(1, "ARRAY厂计划");
//        sheet_name.put(2, "EVEN厂计划");
//        sheet_name.put(3, "TPOT厂计划 ");
//        sheet_name.put(4, "EAC厂计划");
//        sheet_name.put(5, "MODULE厂计划");
//        String cur_sheet_name = "";
//        Gantt_Info gantt_info = new Gantt_Info();
//        int cur_row = 0;
//        for (int i = 1; i <= 5; i++) {
//            cur_sheet_name = sheet_name.get(i);
//            cur_row = PoiExcelTest.readrowNum(filelName,cur_sheet_name);
////            if (i == 2)
////                tmp_row = 9;
////            else if (i == 4)
////                tmp_row = 11;
//            for (int row = 1; row <= cur_row; row++) {
//                for (int col = 0; col < 7; col++) {
//                    if (col == 0) {
//                        gantt_info.setFactory_type(PoiExcelTest.readExcelData(filelName, cur_sheet_name, row, col));
//                    } else if (col == 1) {
//                        gantt_info.setLabel(PoiExcelTest.readExcelData(filelName, cur_sheet_name, row, col));
//                    } else if (col == 2) {
//                        gantt_info.setDepartment(PoiExcelTest.readExcelData(filelName, cur_sheet_name, row, col));
//                    } else if (col == 3) {
//                        gantt_info.setCustomer(PoiExcelTest.readExcelData(filelName, cur_sheet_name, row, col));
//                    } else if (col == 4) {
//                        gantt_info.setOutput_no(PoiExcelTest.readExcelData(filelName, cur_sheet_name, row, col));
//                    } else if (col == 5) {
//                        gantt_info.setTotal(PoiExcelTest.readExcelData(filelName, cur_sheet_name, row, col));
//                    } else if (col == 6) {
//                        gantt_info.setIN_OUTPUT(PoiExcelTest.readExcelData(filelName, cur_sheet_name, row, col));
//                    }
//                }
//                ganttMapper.insert_info(gantt_info);
//            }
//
//        }
//    }
//
//    public void getDetail() throws Exception {
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
//        DateFormat format=new SimpleDateFormat("yyyy/MM/dd");
//        for(int i = 1; i <=5; i++){
//            cur_sheet_name = sheet_name.get(i);
//            cur_row = PoiExcelTest.readrowNum(filelName,cur_sheet_name);
//            cur_col = PoiExcelTest.readcolNum(filelName,cur_sheet_name);
//            for(int row = 1; row < cur_row; row++){
//                gantt_detail.setLabel(PoiExcelTest.readExcelData(filelName,cur_sheet_name,row,1));
//                gantt_detail.setIN_OUTPUT(PoiExcelTest.readExcelData(filelName,cur_sheet_name,row,6));
//                for(int col = 7; col <cur_col; col++){
//                    String time_Str = PoiExcelTest.readExcelData(filelName,cur_sheet_name,0,col);
//                    Date time_date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.valueOf(time_Str));
//                    gantt_detail.setUse_time(format.format(time_date));
//                    if(PoiExcelTest.readExcelData(filelName,cur_sheet_name,row,col).equals(""))
//                        gantt_detail.setUse_amount("0");
//                    else
//                        gantt_detail.setUse_amount(PoiExcelTest.readExcelData(filelName,cur_sheet_name,row,col));
////                    System.out.println(gantt_detail);
//                    ganttMapper.insert_detail(gantt_detail);
//                }
//            }
//        }
//        SimpleDateFormat sdf = new SimpleDateFormat();// 格式化时间
//        sdf.applyPattern("yyyy-MM-dd HH:mm:ss a");// a为am/pm的标记
//        Date date = new Date();// 获取当前时间
//        System.out.println("现在时间：" + sdf.format(date)); // 输出已经格式化的现在时间（24小时制）
//
//    }
//
//
//}
