package com.tianma.yunying.controller;

import com.github.pagehelper.PageInfo;
import com.tianma.yunying.entity.FileInfo;
import com.tianma.yunying.entity.MailInfo;
import com.tianma.yunying.entity.View_Mail;
import com.tianma.yunying.service.InfoService;
import org.apache.poi.ss.usermodel.IconMultiStateFormatting;
import org.apache.tomcat.util.http.fileupload.IOUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.web.bind.annotation.*;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.List;

@RestController
@RequestMapping("/api/file")
public class InfoController {

    @Autowired
    InfoService infoService;
    @Value("${myaddress.upload_excel_path}")
    private String upload_excel_path;

    @CrossOrigin
    @RequestMapping(value = "/all/{pageCode}/{pageSize}/{type}",method = RequestMethod.GET)
    //分页
    public PageInfo<FileInfo> getAllInfo(@PathVariable(value = "pageCode") int pageCode, @PathVariable(value = "pageSize") int pageSize,@PathVariable(value = "type") String role) {
        System.out.println(pageCode+"...."+pageSize);
        PageInfo<FileInfo> pageInfo = infoService.getAllInfo(pageCode, pageSize,role);
        System.out.println(pageInfo);
        System.out.println(role);
        return pageInfo;
    }
    @RequestMapping(value = "/find/{pageCode}/{pageSize}", method = RequestMethod.POST)
    public PageInfo<FileInfo> getFind(@RequestBody FileInfo info, @PathVariable(value = "pageCode") int pageCode, @PathVariable(value = "pageSize") int pageSize) {
        PageInfo<FileInfo> pageInfo = infoService.getFindInfo(pageCode, pageSize,info.getFilename(),info.getImport_user(),info.getRole(),info.getImport_time());
        return  pageInfo;
    }

    @RequestMapping(value = "/findmail/{pageCode}/{pageSize}", method = RequestMethod.POST)
    public PageInfo<View_Mail> getFindMail(@RequestBody View_Mail info, @PathVariable(value = "pageCode") int pageCode, @PathVariable(value = "pageSize") int pageSize) {
        System.out.println("1111");
        System.out.println(info);
        PageInfo<View_Mail> pageInfo = infoService.getFindMailInfo(pageCode, pageSize,info.getTitle(),info.getRole());
        return  pageInfo;
    }
    @RequestMapping(value = "/getAddMail/{pageCode}/{pageSize}/{role}", method = RequestMethod.GET)
    public PageInfo<MailInfo> getAddMail(@PathVariable(value = "pageCode") int pageCode, @PathVariable(value = "pageSize") int pageSize,@PathVariable(value = "role") String role){
        PageInfo<MailInfo> pageInfo = infoService.getAddMail(pageCode, pageSize,role);
        return pageInfo;
    }
    @RequestMapping(value = "/getCCMail", method = RequestMethod.POST)
    public List<MailInfo> getCCMail(String role){
        return infoService.getCCMail(role);
    }

    @RequestMapping(value = "/downloadExcel", method = RequestMethod.GET)
    public void downloadExcel(HttpServletResponse response, String fileName,String role) throws IOException {

        OutputStream os = null;
        InputStream is= null;
        try {
            // 取得输出流
            os = response.getOutputStream();
            // 清空输出流
            response.reset();
            response.setContentType("application/x-download;charset=GBK");
            response.setHeader("Content-Disposition", "attachment;filename="+ new String(fileName.getBytes("utf-8"), "iso-8859-1"));
            //读取流
            File f = new File(upload_excel_path+role+"\\\\"+fileName);

            is = new FileInputStream(f);
            if (is == null) {
//                logger.error("下载附件失败，请检查文件“" + fileName + "”是否存在");
//                return ResultUtil.error("下载附件失败，请检查文件“" + fileName + "”是否存在");
                System.out.println("下载附件失败，请检查文件“" + fileName + "”是否存在");
            }
            //复制
            IOUtils.copy(is, response.getOutputStream());
            response.getOutputStream().flush();
        } catch (IOException e) {
//            return ResultUtil.error("下载附件失败,error:"+e.getMessage());
        }
        //文件的关闭放在finally中
        finally
        {
            try {
                if (is != null) {
                    is.close();
                }
            } catch (IOException e) {
//                logger.error(ExceptionUtils.getFullStackTrace(e));
            }
            try {
                if (os != null) {
                    os.close();
                }
            } catch (IOException e) {
//                logger.error(ExceptionUtils.getFullStackTrace(e));
            }
        }
    }


    @RequestMapping(value = "/allMail/{pageCode}/{pageSize}/{type}", method = RequestMethod.GET)
    public PageInfo<View_Mail> getAllMail(@PathVariable(value = "pageCode") int pageCode, @PathVariable(value = "pageSize") int pageSize,@PathVariable(value = "type") String role){
        System.out.println("1111111111111111111");
        PageInfo<View_Mail> pageInfo = infoService.getAllMail(pageCode, pageSize,role);
        return pageInfo;
    }


}
