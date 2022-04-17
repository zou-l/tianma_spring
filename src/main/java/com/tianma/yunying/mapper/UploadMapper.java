package com.tianma.yunying.mapper;

import com.tianma.yunying.entity.FileInfo;
import com.tianma.yunying.entity.Result;
import com.tianma.yunying.entity.View_Mail;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

//@Repository
@Mapper
public interface UploadMapper {
    int insert(FileInfo record);
    Boolean getFileByname(String filename);
    Boolean getMailByname(View_Mail view_mail);
    void delete(FileInfo record);
    void deleteMail(View_Mail view_mail);
    void insertMailInfo(@Param("title")String title, @Param("content")String content,@Param("upload_time")String upload_time, @Param("role")String role,@Param("filename")String filename);

}
