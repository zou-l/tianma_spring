package com.tianma.yunying.controller;

import com.tianma.yunying.entity.Exam_Histroy;
import com.tianma.yunying.entity.Exam_Info;
import com.tianma.yunying.entity.Exam_User;
import com.tianma.yunying.entity.Result;
import com.tianma.yunying.service.ExamService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;

@RestController
public class ExamController {
    @Autowired
    ExamService examService;

    @RequestMapping("/exam/getUserInfo")
    public Exam_User getUserInfo(String id){
        System.out.println(id);
        return examService.getUserInfo(id);
    }

    @RequestMapping("/exam/insertExamInfo")
    public Result insertExamInfo(Exam_Info exam_info){
        return examService.insertExamInfo(exam_info);
    }

    @RequestMapping("/exam/getHistory")
    public List<Exam_Histroy> getHistory(String exam_id){
        return examService.getHistory(exam_id);
    }
}
