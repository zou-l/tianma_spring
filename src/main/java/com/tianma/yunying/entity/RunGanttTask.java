package com.tianma.yunying.entity;

import lombok.Data;

@Data
public class RunGanttTask {
    private String id;
    private String start_date;
    private String end_date_text;
    private String factory_type;
    private String text;
    private int duration;
    private String parent;
    private String color;
    private Double use_amount;
    private String render;
    private String open;
    private String desc;
    private int number;
    private String pilot;
}
