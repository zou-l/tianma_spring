package com.tianma.yunying.mapper;

import com.tianma.yunying.entity.Exam_Histroy;
import com.tianma.yunying.entity.Exam_Info;
import com.tianma.yunying.entity.Exam_User;

import java.util.List;

public interface ExamMapper {
    Exam_User getUserInfo(String id);
    List<Exam_Histroy> getHistory(String exam_id);
    void insertExamInfo(Exam_Info exam_info);
    void insertExamHistory(Exam_Histroy exam_histroyo);
}
