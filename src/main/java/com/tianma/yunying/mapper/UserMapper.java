package com.tianma.yunying.mapper;

import com.tianma.yunying.entity.User;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;
import org.springframework.stereotype.Repository;


//@Repository
@Mapper
public interface UserMapper {
    User getUserByName(String username);

    Boolean login(@Param("username") String username, @Param("password") String password);

    String getRole(@Param("username") String username, @Param("password") String password);

    void register(String username, String password);

//    Boolean addUser(User user);

}
