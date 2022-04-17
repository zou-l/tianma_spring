package com.tianma.yunying.mapper;

import com.tianma.yunying.entity.FileInfo;
import com.tianma.yunying.entity.MailInfo;
import com.tianma.yunying.entity.View_Mail;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;
import org.springframework.stereotype.Repository;

import java.util.List;

//@Repository
@Mapper
public interface InfoMapper {
    List<FileInfo> getFindInfo(@Param("filename") String filename, @Param("import_user") String import_user, @Param("role") String role,@Param("import_time") String import_time);
    List<FileInfo> getAllInfo(String role);
    List<View_Mail> getAllMail(String role);
    List<View_Mail>getFindMailInfo(@Param("title") String tile,@Param("role") String role);
    List<MailInfo> getAddMail(@Param("facotry") String role);
    List<MailInfo> getCCMail(@Param("facotry") String role);
}
