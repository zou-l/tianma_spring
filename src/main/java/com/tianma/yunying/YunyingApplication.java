package com.tianma.yunying;

import org.mybatis.spring.annotation.MapperScan;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@MapperScan("com.tianma.yunying.mapper")
@SpringBootApplication
public class YunyingApplication {

    public static void main(String[] args) {
        SpringApplication.run(YunyingApplication.class, args);
    }
}
