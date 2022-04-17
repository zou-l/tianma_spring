package com.tianma.yunying.service;

import com.tianma.yunying.entity.Exam_Histroy;
import com.tianma.yunying.entity.Exam_Info;
import com.tianma.yunying.entity.Exam_User;
import com.tianma.yunying.entity.Result;
import com.tianma.yunying.mapper.ExamMapper;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

@Service
public class ExamService {
    @Autowired
    ExamMapper examMapper;
    public Exam_User getUserInfo(String id){
        return  examMapper.getUserInfo(id);
    }
    public List<Exam_Histroy> getHistory(String exam_id){
        return  examMapper.getHistory(exam_id);
    }
    public Result insertExamInfo(Exam_Info exam_info){
        DateFormat format=new SimpleDateFormat("yyyy/MM/dd/HH:MM");
        examMapper.insertExamInfo(exam_info);
        Exam_Histroy exam_histroy = new Exam_Histroy();
        exam_histroy.setExam_id(exam_info.getExam_id());
        exam_histroy.setDetail_desc(exam_info.getContent1());
        exam_histroy.setDetail_task("测试");
        exam_histroy.setId(getUserInfo(exam_info.getId()).getId());
        exam_histroy.setDetail_time(String.valueOf(format.format(new Date().getTime())));
        if(exam_info.getAgree().equals("0"))
            exam_histroy.setDetail_result("驳回");
        else
            exam_histroy.setDetail_result("同意");
        examMapper.insertExamHistory(exam_histroy);
        return new Result("200","成功");
    }
    public Result insertExamHistory(Exam_Histroy exam_histroy){
        examMapper.insertExamHistory(exam_histroy);
        return new Result("200","成功");
    }
}
