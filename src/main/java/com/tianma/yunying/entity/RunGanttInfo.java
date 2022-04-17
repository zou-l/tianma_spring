package com.tianma.yunying.entity;

import lombok.Data;

@Data
public class RunGanttInfo implements Comparable<RunGanttInfo> {
    private String factory_type;
    private String pilot;
    private String department;
    private String customer;
    private String product_number;
    private String desc;
    private String target;
    private Double number;
    private String input_time;
    private String output_time;
    private Double input_amount;
    private Double output_amount;
    private String yield;
    private String cycle;
    private String bank;
    private int duration;
    @Override
    public int compareTo(RunGanttInfo stu) {
        if(this.number == null || stu.getNumber() == null)
            return 1;
        if (this.number <stu.getNumber()) {
            return 1;
        } else {
            return -1;
        }
    }
//    @Override
//    public int compareTo(RunGanttInfo stu) {
//        if (this.input_amount > stu.getInput_amount()) {
//            return 1;
//        } else {
//            return -1;
//        }
//    }
}
