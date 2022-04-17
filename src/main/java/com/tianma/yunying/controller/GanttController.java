package com.tianma.yunying.controller;

import com.tianma.yunying.entity.*;
import com.tianma.yunying.service.GanttTaskService;
import org.json.JSONException;
import org.json.JSONObject;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.IOException;
import java.text.ParseException;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;

@RestController
//@ResponseBody
@RequestMapping("/api/Gantt")
public class GanttController {

    @Autowired
    GanttTaskService ganttTaskService;
    @Value("${myaddress.upload_excel_path}")
    private String upload_excel_path;

    //
//    @RequestMapping(value = "/info", method = RequestMethod.POST)
//    public Result getGanttInfo() throws Exception {
//        ganttService.getInfo();
//        return new Result("200", "success");
//    }
//
//    @RequestMapping(value = "/detail", method = RequestMethod.POST)
//    public Result getGanttDetail() throws Exception {
//        ganttService.getDetail();
//        return new Result("200", "success");
//    }

    @RequestMapping(value = "/task", method = RequestMethod.POST)
    public List<GanttTask> getGanttTask(int type) {
        List<GanttTask> res = ganttTaskService.getTask(type);
        return res;
    }
    @RequestMapping(value = "/insert_task", method = RequestMethod.POST)
    public Result InsertGanttTask(MultipartFile file) throws Exception {
//        String filePath = upload_excel_path + file.getOriginalFilename();
        String filePath = upload_excel_path + "gantt.xlsx";
        File savedFile = new File(filePath);
        if (file != null) {
            try {
                boolean isCreateSuccess = savedFile.createNewFile();
                // 是否创建文件成功
                if (isCreateSuccess) {
                    //将文件写入
                    file.transferTo(savedFile);
//                    Result result = ganttTaskService.new_task(filePath);
                    Result result = ganttTaskService.InsertTask_new(filePath);
//                    Result result1 = ganttTaskService.setMacroGantt(filePath,-1);
//                    Result result1 = ganttTaskService.setMacroGantt();
//                    savedFile.delete();
                    return result;
                } else {
                    savedFile.delete();
                    return new Result("500", "请检查文件格式");
                }
            } catch (Exception e) {
                savedFile.delete();
                e.printStackTrace();
                return new Result("500", "请检查文件格式");
            }
        } else {
            System.out.println("文件是空的");
            savedFile.delete();
            return new Result("500", "请检查文件");
        }
    }
//        return  ganttTaskService.new_task_real();

    @RequestMapping(value = "/insert_task_real", method = RequestMethod.POST)
    public Result InsertGanttTask_Real(MultipartFile file) throws Exception {
        String filePath = upload_excel_path + file.getOriginalFilename();
        File savedFile = new File(filePath);
        if (file != null) {
            try {
                boolean isCreateSuccess = savedFile.createNewFile();
                // 是否创建文件成功
                if (isCreateSuccess) {
                    //将文件写入
                    file.transferTo(savedFile);
                    Result result = ganttTaskService.new_task_real(filePath);
                    savedFile.delete();
                    return result;
                } else {
                    savedFile.delete();
                    return new Result("500", "请检查文件格式");
                }
            } catch (Exception e) {
                e.printStackTrace();
                savedFile.delete();
                return new Result("500", "请检查文件格式");
            }
        } else {
            System.out.println("文件是空的");
            savedFile.delete();
            return new Result("500", "请检查文件");
        }
    }


    @RequestMapping(value = "/runtask", method = RequestMethod.POST)
    public List<RunGanttTask> getRunGanttTask() {
        List<RunGanttTask> res = ganttTaskService.getRunTask();
        return res;
    }

    @RequestMapping(value = "/macrotask", method = RequestMethod.POST)
    public List<MacroGantt> getMacroGanttTask() {
        List<MacroGantt> res = ganttTaskService.getMacroTask();
        return res;
    }


    @RequestMapping(value = "/insert_runtask", method = RequestMethod.POST)
    public Result InsertRunGanttTask(MultipartFile file) throws Exception {
//        return ganttTaskService.InsertRunTask("test");
        String filePath = upload_excel_path + file.getOriginalFilename();
        File savedFile = new File(filePath);
        if (file != null) {
            try {
                boolean isCreateSuccess = savedFile.createNewFile();
                // 是否创建文件成功
                if (isCreateSuccess) {
                    //将文件写入
                    file.transferTo(savedFile);
//                    Result result = ganttTaskService.new_task(filePath);
                    Result result = ganttTaskService.InsertRunTask(filePath);
                    savedFile.delete();
                    return result;
                } else {
                    savedFile.delete();
                    return new Result("500", "请检查文件格式");
                }
            } catch (Exception e) {
                savedFile.delete();
                e.printStackTrace();
                return new Result("500", "请检查文件格式");
            }
        } else {
            System.out.println("文件是空的");
            savedFile.delete();
            return new Result("500", "请检查文件");
        }

    }

    @RequestMapping(value = "/getAllDepart", method = RequestMethod.POST)
    public List<String> getAllDepart(){
        return ganttTaskService.getAllDepart();
    }
    @RequestMapping(value = "/getAllProject", method = RequestMethod.POST)
    public List<String> getAllProject(){
        return ganttTaskService.getAllProject();
    }
    @RequestMapping(value = "/getAllPilot", method = RequestMethod.POST)
    public List<String> getAllPilot(String tablename){
        return ganttTaskService.getAllPilot(tablename);
    }
    @RequestMapping(value = "/getAllTarget", method = RequestMethod.POST)
    public List<String> getAllTarget(){
        return ganttTaskService.getAllTarget();
    }
    @RequestMapping(value = "/runtaskByPilot", method = RequestMethod.POST)
    public List<RunGanttTask> getRunGanttTaskByPilot(String pilot) {
        List<RunGanttTask> res = ganttTaskService.getRunTaskByPilot(pilot);
        return res;
    }
    @RequestMapping(value = "/taskByPilot", method = RequestMethod.POST)
    public List<GanttTask> getGanttTaskByPilot(String pilot) {
        List<GanttTask> res = ganttTaskService.getTaskByPilot(pilot);
        return res;
    }

    @RequestMapping(value = "/getCapacity", method = RequestMethod.POST)
    public List<GanttCapacity> getCapacity(){
        return ganttTaskService.getCapacity();
    }

    @RequestMapping(value = "/getMacroStatus", method = RequestMethod.POST)
    public MacroGantt_Status getMacroStatus(){
        return ganttTaskService.getMacroStatus();
    }
    @RequestMapping(value = "/updateMacro", method = RequestMethod.POST)
    public Result toupdateMacro(MacroGantt_Status macroGantt_status) throws Exception {
        if(macroGantt_status.getDepart().equals("")||macroGantt_status.getDepart().equals("all"))
            macroGantt_status.setDepart("-1");
        if(macroGantt_status.getPilot().equals("")||macroGantt_status.getPilot().equals("all"))
            macroGantt_status.setPilot("-1");
        if(macroGantt_status.getProject().equals("")||macroGantt_status.getProject().equals("all"))
            macroGantt_status.setProject("-1");
        if(macroGantt_status.getTarget().equals("")||macroGantt_status.getTarget().equals("all"))
            macroGantt_status.setTarget("-1");
        return ganttTaskService.setMacroGantt(macroGantt_status);
    }


    @RequestMapping(value = "/updateCapacity", method = RequestMethod.POST)
    public void updateCapacity(GanttCapacity ganttCapacity){
        ganttTaskService.updateCapacity(ganttCapacity);
        System.out.println(ganttCapacity);
    }

    @RequestMapping(value = "/runGanttdownload", method = RequestMethod.POST)
    public void runGantt_download() throws IOException, ParseException {
        ganttTaskService.runGanttdownload();
    }
}
