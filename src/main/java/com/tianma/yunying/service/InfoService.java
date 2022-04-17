package com.tianma.yunying.service;

import com.github.pagehelper.PageHelper;
import com.github.pagehelper.PageInfo;
import com.tianma.yunying.entity.FileInfo;
import com.tianma.yunying.entity.MailInfo;
import com.tianma.yunying.entity.View_Mail;
import com.tianma.yunying.mapper.InfoMapper;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.List;

@Service
public class InfoService {
    @Autowired
    InfoMapper infoMapper;

    public PageInfo<FileInfo> getAllInfo(int pageCode, int pageSize,String role) {

        PageHelper.startPage(pageCode,pageSize);
        List<FileInfo> classInfos = infoMapper.getAllInfo(role);
        PageInfo<FileInfo> test =  new PageInfo<FileInfo>(classInfos);
        PageHelper.clearPage();
        return test;
    }

    public PageInfo<FileInfo>getFindInfo(int pageCode, int pageSize,String filename,String import_user,String role,String import_time) {
        int number = 0;
        PageHelper.startPage(pageCode,pageSize);
        List<FileInfo> classInfos = infoMapper.getFindInfo(filename,import_user,role,import_time);
        PageInfo<FileInfo> test =  new PageInfo<FileInfo>(classInfos);
        return test;
    }
    public PageInfo<View_Mail> getAllMail(int pageCode, int pageSize,String role){
        PageHelper.startPage(pageCode,pageSize);
        List<View_Mail> classInfos = infoMapper.getAllMail(role);
        PageInfo<View_Mail> mailInfo =  new PageInfo<View_Mail>(classInfos);
        return mailInfo;
    }

    public PageInfo<View_Mail> getFindMailInfo(int pageCode, int pageSize,String title, String role){
        PageHelper.startPage(pageCode,pageSize);
        List<View_Mail> classInfos = infoMapper.getFindMailInfo(title,role);
        PageInfo<View_Mail> mailInfo =  new PageInfo<View_Mail>(classInfos);
        return mailInfo;
    }
    public PageInfo<MailInfo> getAddMail(int pageCode, int pageSize, String role){
        PageHelper.startPage(pageCode,pageSize);
        List<MailInfo> classInfos = infoMapper.getAddMail(role);
        PageInfo<MailInfo> mailInfo =  new PageInfo<MailInfo>(classInfos);
        return mailInfo;
    }
    public List<MailInfo> getAddAllMail(String role){
        return infoMapper.getAddMail(role);
    }
    public List<MailInfo> getCCMail(String role){
        return infoMapper.getCCMail(role);
    }

}
