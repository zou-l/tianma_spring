package com.tianma.yunying.controller;

import com.tianma.yunying.entity.Result;
import com.tianma.yunying.service.UserService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;

@RequestMapping("/api")
@RestController
public class UserController {
    @Autowired
    private UserService userService;

//    @RequestMapping("getUser/{username}")
//    public String  getUser(@PathVariable String username){
//        return userService.getUserByName(username).toString();
//    }

    @RequestMapping("/login")
    public Result login(String username, String password){
        return userService.login(username,password);
    }

    @RequestMapping("/register")
    public Result register(String username, String password){
        return userService.register(username,password);
    }
    @RequestMapping(value = "/getUser",method = RequestMethod.POST)
    public String  getUser(String username){
        System.out.println("^^^^^^^^^^^^^^");
        System.out.println(username);
        return userService.getUserByName(username).toString();
    }
}
