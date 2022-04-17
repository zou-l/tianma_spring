package com.tianma.yunying.mapper;

import com.tianma.yunying.entity.*;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.List;

@Mapper
public interface GanttTaskMapper {
    List<GanttTask> getTask_plan();
    List<GanttTask> getTask_real();
    List<GanttTask> getTask();
    List<RunGanttTask> getRunTask();
    List<RunGanttTask> getRunTaskByPilot(String pilot);
    MacroGantt_Status getMacroStatus();
    List<MacroGantt> getMacroTask();
    List<GanttTask> getTaskByPilot(String pilot);
    List<String> getAllDepart();
    List<String> getAllProject();
    List<String> getAllPilot(@Param("tablename") String tablename);
    List<String> getAllTarget();
    List<GanttCapacity> getCapacity();
    List<RunGanttInfo>getRunGanttInfo();
    void updateCapacity(GanttCapacity ganttCapacity);
    void updatePlanEnd(@Param("id") String id, @Param("end_date_text")String end_date_text);
    void deleteMacroStatus();
    void deleteTask();
    void deleteRunTask();
    void deleteRunGanttInfo();
    void deleteMacroTask();
    void insertTask(GanttTask ganttTask);
    void insertTask_run(RunGanttTask runGanttTask);
    void insertRunGanttInfo(RunGanttInfo runGanttInfo);
    void insertMacroGantt(MacroGantt macroGantt);
    void insertMacroStatus(MacroGantt_Status macroGantt_status);
}
