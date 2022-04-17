package com.tianma.yunying.config;

import org.springframework.boot.SpringBootConfiguration;
import org.springframework.web.servlet.config.annotation.ResourceHandlerRegistry;
import org.springframework.web.servlet.config.annotation.WebMvcConfigurer;

@SpringBootConfiguration
public class MyWebConfigurer implements WebMvcConfigurer {

@Override
public void addResourceHandlers(ResourceHandlerRegistry registry) {
    registry.addResourceHandler("/api/file/image/**").addResourceLocations("file:" + "D:\\yunying\\upload\\picture\\");
    registry.addResourceHandler("/api/mail/**/**").addResourceLocations("file:D:\\yunying\\upload\\mail\\");
    registry.addResourceHandler("/api/plan/**").addResourceLocations("file:D:\\yunying\\upload\\excel\\htm\\");

}

}

