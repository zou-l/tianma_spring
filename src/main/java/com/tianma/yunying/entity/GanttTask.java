package com.tianma.yunying.entity;

import lombok.Data;

@Data
public class GanttTask {
    private String id;
    private String start_date;
    private String end_date_text;
    private String text;
    private int duration;
    private String parent;
    private String color;
    private String use_amount;
    private String render;
    private String open;
    private String desc;
    private String project_name;
    private String gantt_type;
    private int number;
    private String factory_number;
    private String pilot;
    private String code1;
    private String code2;
    private String code3;
    private String code4;
}
