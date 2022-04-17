package com.tianma.yunying.service;

import com.tianma.yunying.entity.FileInfo;
import com.tianma.yunying.entity.Result;
import com.tianma.yunying.entity.View_Mail;
import com.tianma.yunying.mapper.UploadMapper;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

@Service
public class uploadService {
//    @Autowired
//    ResourceRepository repository;
    @Autowired
UploadMapper uploadMapper;

    public Result insertInfo(FileInfo fileInfo){
        uploadMapper.insert(fileInfo);
        return new Result("success","导入成功");
    }
    public Boolean getFileByname(FileInfo fileInfo){
        if(uploadMapper.getFileByname(fileInfo.getFilename())){
            return true;
        }
        else
            return  false;
    }

    public Result deleteInfo(FileInfo fileInfo){
        uploadMapper.delete(fileInfo);
        return new Result("success","删除成功");
    }
    public Boolean getMailByname(View_Mail view_mail){
        if(uploadMapper.getMailByname(view_mail)){
            return true;
        }
        else
            return false;
    }
    public Result deleteMailInfo(View_Mail view_mail){
        uploadMapper.deleteMail(view_mail);
        return new Result("success","删除成功");
    }

    public void insertMailInfo(String title,String content,String upload_time,String role,String filename){
         uploadMapper.insertMailInfo(title,content,upload_time,role,filename);
    }
}
