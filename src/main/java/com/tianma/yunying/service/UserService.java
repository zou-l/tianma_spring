package com.tianma.yunying.service;

import com.tianma.yunying.entity.Result;
import com.tianma.yunying.entity.User;
import com.tianma.yunying.mapper.UserMapper;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

@Service
public class UserService {
    @Autowired
    UserMapper userMapper;

    public User getUserByName(String username){
        return userMapper.getUserByName(username);
    }

    public Result login(String username, String password) {
        if (userMapper.login(username,password) == null){
            String message = "账号密码错误";
            return new Result("error","登录失败");
        }
        else {
            return new Result("success",userMapper.getRole(username,password));
        }
    }

//    public Result test(String a,String b){
//
//    }
    public Result register(String username,String password){

        System.out.println(username);
        User existUser = userMapper.getUserByName(username);
        if(existUser!=null){
            return new Result("error","注册失败");
        }
        else{
            userMapper.register(username,password);
            return new Result("success","注册成功");
        }
    }
}
