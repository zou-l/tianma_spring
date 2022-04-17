package com.tianma.yunying.entity;

import lombok.Data;

@Data
public class Result {
    private String status;
    private String msg;

    public Result(String status, String msg) {
        this.status = status;
        this.msg = msg;
    }
}
