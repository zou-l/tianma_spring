package com.tianma.yunying.mapper;

import com.tianma.yunying.entity.Gantt_Detail;
import com.tianma.yunying.entity.Gantt_Info;
import org.apache.ibatis.annotations.Mapper;

@Mapper
public interface GanttMapper {
    void insert_info(Gantt_Info gantt_info);
    void insert_detail(Gantt_Detail gantt_detail);
}
