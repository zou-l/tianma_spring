package com.tianma.yunying.controller;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.tianma.yunying.entity.FileInfo;
import com.tianma.yunying.entity.MailInfo;
import com.tianma.yunying.entity.Result;
import com.tianma.yunying.entity.View_Mail;
import com.tianma.yunying.service.InfoService;
import com.tianma.yunying.service.uploadService;
import com.tianma.yunying.test.ExcelTest;
import com.tianma.yunying.util.ExcelUtil;
import com.tianma.yunying.util.FileUtil;
import com.tianma.yunying.util.StringUtils;
import com.tianma.yunying.util.ZipUtil;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.apache.tomcat.util.http.fileupload.IOUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Map;

//import com.jacob.activeX.ActiveXComponent;
//import com.jacob.com.Dispatch;
//import org.apache.poi.hssf.util.Region;

/**
 *
 * @author
 *
 */
@RestController
@RequestMapping("/api/file")
public class UploadController {
    @Autowired
    uploadService uploadService;
    @Autowired
    InfoService infoService;
    @Value("${EmailInfo.myEmailAccount}")
    private String myEmailAccount;
    @Value("${EmailInfo.myEmailPassword}")
    private String myEmailPassword;
    // 发件人邮箱的 SMTP 服务器地址, 必须准确, 不同邮件服务器地址不同, 一般(只是一般, 绝非绝对)格式为: smtp.xxx.com
    @Value("${EmailInfo.myEmailSMTPHost}")
    private String myEmailSMTPHost;
    @Value("${myaddress.ip}")
    private String myhost;
    @Value("${myaddress.upload_excel_path}")
    private String upload_excel_path;
    @Value("${myaddress.upload_picture_path}")
    private String upload_picture_path;
    @Value("${myaddress.upload_mail_path}")
    private String upload_mail_path;
    // 收件人邮箱（替换为自己知道的有效邮箱）
//    public static String receiveMailAccount = "1754687268@qq.com";
    public static String cur_excel_path = "";


    @RequestMapping(value = "/upload", method = RequestMethod.POST)
    @ResponseBody
    Result transfer(MultipartFile file, String import_user, String role) {
        String filePath = upload_excel_path + role + "\\\\" + file.getOriginalFilename();

        if (file != null) {

            File deleteFile = new File(filePath);
            deleteFile.delete();
            try {
//                String filePath = upload_excel_path + role+"\\\\"+file.getOriginalFilename();
                FileInfo tmp = new FileInfo();
                tmp.setFilename(file.getOriginalFilename());
                System.out.println(filePath);
                SimpleDateFormat df = new SimpleDateFormat("yyyyMMdd");
                tmp.setImport_time(df.format(new Date()));
                tmp.setImport_user(import_user);
                if (uploadService.getFileByname(tmp)) {
                    uploadService.deleteInfo(tmp);
                    uploadService.insertInfo(tmp);
                } else {
                    uploadService.insertInfo(tmp);
                }

                File savedFile = new File(filePath);
                boolean isCreateSuccess = savedFile.createNewFile();
                // 是否创建文件成功
                if (isCreateSuccess) {
                    //将文件写入
                    file.transferTo(savedFile);
                    if(role.equals("EVEN")){
                        ExcelTest.OpenExcel(filePath,false);
                        FileUtil.delAllFile("D:\\yunying\\upload\\excel\\htm\\");
                        File file1 = new File("D:\\yunying\\upload\\excel\\htm\\EVEN生产计划.files");
                        file1.delete();
                        ExcelTest.SaveAs("D:\\yunying\\upload\\excel\\htm\\EVEN生产计划.htm");
                        ExcelTest.releaseSource();
//                        return null;
                        return new Result("200","成功");
                    }
                    return new Result("200","成功");
//                    return savedFile;
                }
            } catch (Exception e) {
                e.printStackTrace();
                return new Result("405","失败");
            }
        } else {
            System.out.println("文件是空的");
            return new Result("500","失败");
        }
//        return null;
        return new Result("200","成功");
    }


    @GetMapping("/excelExport")
    public ResponseEntity<byte[]> excel(String role) throws IOException {
//        String filePath = upload_excel_path+"array_"+df.format(new Date())+".xlsx";
//        String filePath = upload_excel_path+role+"\\\\"+role+"_"+df.format(new Date())+".xlsx";

        long cur_time = new Date().getTime();
        Boolean isExist = false;
        int time = 0;
        long tmp = cur_time;
        int rownum_tmp = 1;
        String cur_tmp = "";
        String filePath_tmp = "";
        File file;
        int factoroy_change = 0;
        System.out.println(role);
        if(role.equals("EAC")){
            factoroy_change = 2;
            rownum_tmp = 5;
        }
        while(!isExist && time < 30){
            cur_tmp = StringUtils.timeStamp2Date(tmp);
            System.out.println(cur_tmp);
            filePath_tmp = upload_excel_path+role+"\\\\"+role+"生产计划_"+cur_tmp.replace(".","")+".xlsx";
            file = new File(filePath_tmp);
            if(!file.exists()){
                tmp = cur_time - 86400000;
                cur_time = tmp;
                time += 1;
            }
            else {
                isExist = true;
            }
        }
        cur_excel_path = filePath_tmp;
        String filePath = cur_excel_path;
        System.out.println(filePath_tmp);

//        String filePath = upload_excel_path+"data_"+"2022.02.22.xlsx";
        List<Object> objects3 = ExcelUtil.readMoreThan1000Row(filePath);
        int t_cellnum = objects3.get(2+factoroy_change).toString().replace("[","").replace("]","").split(", ").length;
        System.out.println("11111111111111111111111111111");
        System.out.println(objects3.get(2+factoroy_change).toString().replace("[","").replace("]","").split(", ")[3]);
        int type_number = 0;
        for(int i = 0; i <t_cellnum; i++){
            String tmp_var = objects3.get(2+factoroy_change).toString().replace("[","").replace("]","").split(", ")[i];
            if(tmp_var.equals("IN/OUT")){
                type_number = i;
            }
        }
        System.out.println(type_number);
        int  t_rownum = objects3.size();
        System.out.println(t_cellnum);
        System.out.println(t_rownum);
        String[] excel_tmp = new String[t_cellnum];
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet();
        XSSFCellStyle style = wb.createCellStyle();

        //设置底边框颜色;
//        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        //设置自定义填充颜色
//        style.setFillForegroundColor(new XSSFColor(new Color(20,13,255)));

        style.setFillForegroundColor(new XSSFColor(new java.awt.Color(135,206,250), new DefaultIndexedColorMap()));
        XSSFFont font = wb.createFont();
        font.setFontName("黑体");
        font.setFontHeightInPoints((short) 17);
        style.setFont(font);
        DateFormat format=new SimpleDateFormat("MM/dd");

        //设置单元格合并
//        for(int rownum1 = 1; rownum1 < t_rownum; rownum1++){
            for(int rownum1 = rownum_tmp; rownum1 < t_rownum - 2-factoroy_change; rownum1++){
            for(int colnum = 0; colnum < type_number; colnum++){
                    sheet.addMergedRegion(new CellRangeAddress(rownum1,rownum1+1,colnum,colnum));
//                }

            }
            rownum1++;
        }

        for (int rownum = 0; rownum < t_rownum - 2 + factoroy_change; rownum++) {
            XSSFRow hssfRow = sheet.createRow(rownum);
            if(rownum > 0){
                for(int i = 0; i < t_cellnum; i++){
                    excel_tmp[i] = " ";
                }
                for(int i = 0 ; i < t_cellnum; i++){
                    try {
                        excel_tmp[i] = objects3.get(rownum + 2 - factoroy_change).toString().replace("[","").replace("]","").replace("null","").split(", ")[i];
                    } catch (ArrayIndexOutOfBoundsException e){
//                        System.out.println("空值: " + e);
                    }

                }
            }
            for (int cellnum = 0; cellnum < t_cellnum; cellnum++) {
                XSSFCell cell = hssfRow.createCell((short) cellnum);
//                cell.setCellValue(objects3.get(rownum).toString().replace("[","").replace("]","").split(", ")[cellnum]);
                if(rownum == 0){
                    if(cellnum > type_number){
                        String time_Str = objects3.get(rownum+2+factoroy_change).toString().replace("[","").replace("]","").split(", ")[cellnum];
                        if(StringUtils.isNumeric(time_Str) && !time_Str.equals("")){
                            Date time_date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.valueOf(time_Str));
                            cell.setCellValue(format.format(time_date));
                        }
                        else {
                            cell.setCellValue(time_Str);
                        }
                    }
                    else{
                        String temp_var = objects3.get(rownum+2-factoroy_change).toString().replace("[","").replace("]","").split(", ")[cellnum];
                        if(temp_var.equals("null")){
                            temp_var = "";
                        }
                            cell.setCellValue(temp_var);
                    }
                    cell.setCellStyle(style);
                }
                else if(rownum == rownum_tmp - 1){
                    if(cellnum > type_number){
                        String time_Str = objects3.get(rownum).toString().replace("[","").replace("]","").split(", ")[cellnum];
                        if(StringUtils.isNumeric(time_Str) && !time_Str.equals("")){
                            Date time_date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.valueOf(time_Str));
                            cell.setCellValue(format.format(time_date));
                        }
                        else {
                            cell.setCellValue(time_Str);
                        }
                    }
                    else{
                        cell.setCellValue(objects3.get(rownum+2-factoroy_change).toString().replace("[","").replace("]","").split(", ")[cellnum]);
                    }
                    cell.setCellStyle(style);
                }
                else {
                    if(cellnum < type_number){
                        String time_Str = excel_tmp[cellnum];
                        if(StringUtils.isNumeric(excel_tmp[cellnum])){
                            if (!excel_tmp[cellnum].isEmpty() && Double.parseDouble(excel_tmp[cellnum]) > 10000) {
                                Date time_date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.valueOf(time_Str));
                                cell.setCellValue(format.format(time_date));
                            } else
                                cell.setCellValue(excel_tmp[cellnum]);
                        }
                        else {
                            cell.setCellValue(excel_tmp[cellnum]);
                        }
                    }
                       else {
                           if(StringUtils.isNumeric(excel_tmp[cellnum])&&!excel_tmp[cellnum].isEmpty()){
                               cell.setCellValue(Math.round(Double.parseDouble(excel_tmp[cellnum])));
                           }

                        }
                }
            }
        }

        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        try {
            wb.write(outputStream);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            outputStream.close();
        }
        HttpHeaders httpHeaders = new HttpHeaders();
//        String fileName = new String("测试.xls".getBytes("UTF-8"), "iso-8859-1");
//        FileOutputStream fileOut = new FileOutputStream(
//                "测试.xlsx");
//        wb.write(fileOut);
//        fileOut.close();
//        httpHeaders.setContentDispositionFormData("attachment", fileName);
//        httpHeaders.setContentType(MediaType.APPLICATION_OCTET_STREAM);
        ResponseEntity<byte[]> filebyte = new ResponseEntity<byte[]>(outputStream.toByteArray(), httpHeaders, HttpStatus.CREATED);
        try {
            outputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            outputStream.close();
        }
        return filebyte;
    }

    @CrossOrigin
    @PostMapping("/image")
    public String coversUpload(MultipartFile file,String role) throws Exception {
        File imageFolder = new File(upload_picture_path+role+"\\");
//        File f = new File(imageFolder, file.getOriginalFilename());
        File f = new File(imageFolder, "甘特图.png");
        if (!f.getParentFile().exists())
            f.getParentFile().mkdirs();
        try {
            file.transferTo(f);
//            String imgURL = "http://"+myhost+":8081/api/file/image/"+role+"/" + f.getName();
            String imgURL = "http://"+myhost+":8081/api/file/image/"+role+"/" + f.getName();
            System.out.println(imgURL);
            return imgURL;
        } catch (IOException e) {
            e.printStackTrace();
            return "";
        }
    }

    @RequestMapping("/sendNoFile")
    public Result send_without_file(String content,String role,String title) throws Exception {
        List<MailInfo> add = infoService.getAddAllMail(role);
        List<MailInfo> cc = infoService.getCCMail(role);
        String cc_str = "";
        for(int i = 0; i < cc.size();i++){
            cc_str += cc.get(i).getMail()+";";
        }
        System.out.println(cc_str);
        DateFormat format=new SimpleDateFormat("yyyy/MM/dd");
        long cur_time = new Date().getTime();
        long cur_time2 = new Date().getTime();
        Boolean isExist = false;
        int time = 0;
        long tmp = cur_time;
        int rownum_tmp = 1;
        String cur_tmp = "";
        String filePath_tmp = "";
        File file;
        ActiveXComponent axOutlook = null;
        System.out.println(title);
        try{
//            System.out.println(role);
            axOutlook = new ActiveXComponent("Outlook.Application");
            Dispatch mailItem = Dispatch.call(axOutlook, "CreateItem", 0).getDispatch();
            Dispatch inspector = Dispatch.get(mailItem, "GetInspector").getDispatch();
            //收件人
            Dispatch recipients = Dispatch.call(mailItem, "Recipients").getDispatch();

//            Dispatch.call(recipients, "Add" , "1754687268@qq.com"); //添加收件人
            for(int i = 0; i < add.size();i++){
                Dispatch.call(recipients, "Add" , add.get(i).getMail()); //添加收件人
                System.out.println(add.get(i).getMail());
            }
//            Dispatch.call(recipients, "Add" , "lu_zou@tianma.cn"); //添加收件人
//            if(role.equals("ARRAY")) {
//                Dispatch.put(mailItem, "CC", "lu_zou@tianma.cn;guixia_ye@tianma.cn"); //抄送
//            }
//            if(role.equals("EVEN")) {
//                Dispatch.put(mailItem, "CC", "lu_zou@tianma.cn;yisen_chen@tianma.cn"); //抄送
//            }
//            if(role.equals("TPOT")) {
//                Dispatch.put(mailItem, "CC", "lu_zou@tianma.cn;yili_lin@tianma.cn"); //抄送
//            }
            if(role.equals("EAC")){
//                Dispatch.call(recipients,"Add","tm18_oled_yunying_jihua@tianma.cn");
//                Dispatch.call(recipients,"Add","oled_xmgl2_npm@tianma.cn");
//                Dispatch.call(recipients,"Add","yuanxing_deng@tianma.cn");
//                Dispatch.call(recipients,"Add","chunfang_zhu@tianma.cn");
//                Dispatch.call(recipients,"Add","weijie_cai@tianma.cn");
//                Dispatch.call(recipients,"Add","tao_zhang13@tianma.cn");
//                Dispatch.call(recipients,"Add","liangjian_ye@tianma.cn");
//                Dispatch.call(recipients,"Add","zhaofan_wang@tianma.cn");
//                Dispatch.call(recipients,"Add","kun_fan@tianma.cn");
//                Dispatch.call(recipients,"Add","wenhao_zhu@tianma.cn");
//                Dispatch.call(recipients,"Add","guosong_yu@tianma.cn");
//                Dispatch.call(recipients,"Add","wenling_xu@tianma.cn");
//                Dispatch.call(recipients,"Add","xunfeng_xu@tianma.cn");
//                Dispatch.call(recipients,"Add","lan_li3@tianma.cn");
//                Dispatch.call(recipients,"Add","xia_yuan2@tianma.cn");
//                Dispatch.call(recipients,"Add","yangfeng_shao1@tianma.cn");
//                Dispatch.call(recipients,"Add","ying_wei@tianma.cn");
//                Dispatch.call(recipients,"Add","yuqun_hu@tianma.cn");
//                Dispatch.call(recipients,"Add","heying_mao@tianma.cn");
//                Dispatch.call(recipients,"Add","joe_hahn@tianma.cn");
//                Dispatch.call(recipients,"Add","qunteng_zheng@tianma.cn");
//                Dispatch.call(recipients,"Add","xiaolian_zhou@tianma.cn");
//                Dispatch.call(recipients,"Add","ruiying_gao@tianma.cn");
//                Dispatch.call(recipients,"Add","dongai_shen@tianma.cn");
//                Dispatch.call(recipients,"Add","guangyan_jiang@tianma.cn");
//                Dispatch.call(recipients,"Add","dapeng_li@tianma.cn");
//                Dispatch.call(recipients,"Add","cunjun_xia@tianma.cn");
//                Dispatch.call(recipients,"Add","shulian_wang@tianma.cn");
//                Dispatch.call(recipients,"Add","jiulong_zhang1@tianma.cn");
//                Dispatch.call(recipients,"Add","xiangxu_meng@tianma.cn");
//                Dispatch.call(recipients,"Add","zhiyuan_wang2@tianma.cn");
//                Dispatch.call(recipients,"Add","lili_xia@tianma.cn");
//                Dispatch.call(recipients,"Add","mengling_zheng@tianma.cn");
//                Dispatch.call(recipients,"Add","binbin_xu@tianma.cn");
//                Dispatch.call(recipients,"Add","xmoled_qa_xcpk@tianma.cn");

//                Dispatch.put(mailItem, "CC", "ling_qiu@tianma.cn"); //抄送
//                Dispatch.put(mailItem, "CC", "yuanyi_zhang@tianma.cn"); //抄送
//                Dispatch.put(mailItem, "CC", "yunfei_qiu@tianma.cn"); //抄送
//                Dispatch.put(mailItem, "CC", "julong_si@tianma.cn"); //抄送
//                Dispatch.put(mailItem, "CC", "david@tianma.cn"); //抄送


            }
            SimpleDateFormat df = new SimpleDateFormat("yyyyMMdd");
            String date = df.format(new Date());
            String filePath = cur_excel_path;
            String tmp_name = "";
//            if(role.equals("EVEN") || filePath.equals("")){
                while(!isExist && time < 60){
                    cur_tmp = StringUtils.timeStamp2Date(tmp);
                    System.out.println(cur_tmp);
                    tmp_name = role+"\\\\"+role+"生产计划_"+cur_tmp.replace(".","")+".xlsx";
                    filePath_tmp = upload_excel_path+tmp_name;
                    file = new File(filePath_tmp);
                    if(!file.exists()){
                        tmp = cur_time - 86400000;
                        cur_time = tmp;
                        time += 1;
                    }
                    else {
                        isExist = true;
                    }
                    cur_excel_path = filePath_tmp;
                    filePath = cur_excel_path;
                    System.out.println(filePath_tmp);
                }
            View_Mail view_mail = new View_Mail();
            view_mail.setTitle(title);
            view_mail.setFilename(role+"生产计划_"+cur_tmp.replace(".","")+".xlsx");
            if (uploadService.getMailByname(view_mail)) {
                uploadService.deleteMailInfo(view_mail);
                uploadService.insertMailInfo(title,content,format.format(cur_time2),role,role+"生产计划_"+cur_tmp.replace(".","")+".xlsx");
            } else {
                uploadService.insertMailInfo(title,content,format.format(cur_time2),role,role+"生产计划_"+cur_tmp.replace(".","")+".xlsx");
            }
//            }

//            System.out.println();
//            Dispatch.put(mailItem, "Subject", "TM18 "+role+"工厂M+1生产计划_"+role+"_"+cur_excel_path.split("_")[1]); //主题

            Dispatch.put(mailItem, "Subject",title); //主题

            Dispatch.put(mailItem, "CC", cc_str); //抄送

//            Dispatch.put(mailItem, "CC", "1754687268@qq.com");
            //Dispatch.put(mailItem, "ReadReceiptRequested", "false");

//            String body = "<html><body>" +
//                    "<div> <p>大家好<br />\n" +
//                    " 这是今天的内容:</p>"
//                    +content+
//                    "</div></body></html>";

            String body = "<html><body><div style=\"white-space:pre-wrap;\" v-html=\"data.quillContent\">"+
//                    "<p><b>Dear All:</b></p>\n" +
//                    "\n" +
//                    "<p> 大家好！附件为TM18 "+role+"工厂M+1生产计划，请查收！</p>\n" +

                    "<p>" + content +"</p>"+
                    "<p>以上，祝好！</p> " +
                    "<p>------------------------</p>" +
                    "<p>运营部_计划分部</p> " +
                    "<img src='"+upload_mail_path+"tianma_logo.png'>" +
                    " <p>厦门天马显示科技有限公司</p>" +
                    " <p>Xiamen Tianma Display Technology Co.,Ltd. </p>" +
                    "<p>福建省厦门市翔安区翔安西路6999号 No.6999,West Xiangan Road,XianganDistrict,Xiamen,China</p>" +
                    "</div></body></html>";
            Date day=new Date();
            SimpleDateFormat tmp_df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
            String body_save = "<html>  <head>\n" +
                    "                <meta http-equiv=\"Content-Type\" content=\"text/html; charset=gbk\" />\n" +
                    "            </head><body><div>"+
//                    "<p><b>Dear All:</b></p>\n" +
//                    "\n" +
//                    "<p> 大家好！附件为TM18 "+role+"工厂M+1生产计划，请查收！</p>\n" +
                    "<p>" + content +"</p>"+
//                    "<img src='"+upload_mail_path +map.get("role")+"\\\\"+file.getOriginalFilename()+"' width = 180% height = 100%>"+

                    "<p><b>发布时间为："+ tmp_df.format(day)+"</b></p>"+
                    "</div></body></html>";
            FileUtil.write(upload_mail_path+role+"\\\\public.html",body_save);

//            String body = map.get("content");
//            String content = body + "<img src='"+mail_filePath+file.getOriginalFilename()+"'>";
            Dispatch.put(mailItem, "HTMLBody", body);

            //附件
            Dispatch attachments = Dispatch.call(mailItem, "Attachments").getDispatch();
//            String filePath = upload_excel_path+role+"\\\\"+role+"_"+df.format(new Date())+".xlsx";
//            String filePath = cur_excel_path;
            System.out.println(filePath);
            Dispatch.call(attachments, "Add" , filePath);
            Dispatch.call(mailItem, "Display");
            Dispatch.call(mailItem, "Send");

            System.out.println("1111111111111111111111111111111111111111111111111111111111111");
            return new Result("200", tmp_name);
        }
        catch (Exception e) {
            System.out.println(("调用outlook失败,无法发送邮件"));
            return new Result("405", "fail");
        }

    }


    @RequestMapping("/send")
    public Result send(@RequestParam Map<String, String> map, @RequestParam("file") MultipartFile file) throws Exception {
        List<MailInfo> add = infoService.getAddAllMail(map.get("role"));
        List<MailInfo> cc = infoService.getCCMail(map.get("role"));
        String cc_str = "";
        for(int i = 0; i < cc.size();i++){
            cc_str += cc.get(i).getMail()+";";
        }
        System.out.println(cc_str);
        DateFormat format=new SimpleDateFormat("yyyy/MM/dd");
        long cur_time = new Date().getTime();
        long cur_time2 = new Date().getTime();
        Boolean isExist = false;
        int time = 0;
        long tmp = cur_time;
        int rownum_tmp = 1;
        String cur_tmp = "";
        String filePath_tmp = "";
        File file2;
        System.out.println("!!!!!!!!!!!!!!!!!!");
        System.out.println(map.get("role"));
        String full_filePath = upload_mail_path +map.get("role")+"\\\\"+file.getOriginalFilename();
        String mailFolder =  upload_mail_path+map.get("role")+"\\\\";
        File f = new File(mailFolder, file.getOriginalFilename());
        if (!f.getParentFile().exists())
            f.getParentFile().mkdirs();
            try {
                file.transferTo(f);
                System.out.println(full_filePath);
                File savedFile = new File(full_filePath);
                boolean isCreateSuccess = savedFile.createNewFile();
                // 是否创建文件成功
                if (isCreateSuccess) {
                    //将文件写入
                    file.transferTo(savedFile);
                }
            } catch (Exception e) {
                e.printStackTrace();
        }
        ActiveXComponent axOutlook = null;
        try{
            axOutlook = new ActiveXComponent("Outlook.Application");
            Dispatch mailItem = Dispatch.call(axOutlook, "CreateItem", 0).getDispatch();
            Dispatch inspector = Dispatch.get(mailItem, "GetInspector").getDispatch();
            //收件人
            Dispatch recipients = Dispatch.call(mailItem, "Recipients").getDispatch();

            for(int i = 0; i < add.size();i++){
                Dispatch.call(recipients, "Add" , add.get(i).getMail()); //添加收件人
                System.out.println(add.get(i).getMail());
            }



            SimpleDateFormat df = new SimpleDateFormat("yyyyMMdd");
            String date = df.format(new Date());
//            Dispatch.put(mailItem, "Subject", "TM18 "+map.get("role")+"工厂M+1生产计划_"+map.get("role")+"_"+cur_excel_path.split("_")[1]); //主题

            Dispatch.put(mailItem, "Subject", map.get("title")); //主题
            Dispatch.put(mailItem, "CC", cc_str); //抄送

//            String body = "<html><body>" +
//                    "<div> <p>大家好<br />\n" +
//                    " 这是今天的内容:</p>"+
//                    map.get("content")+
//                    "<img src='"+upload_mail_path +map.get("role")+"\\\\"+file.getOriginalFilename()+"' width = 180% height = 100%>" +
//                    "</div></body></html>";

            String body = "<html><body><div style=\"white-space:pre-wrap;\" v-html=\"data.quillContent\">"+
                    "<p class=\"ql-editor\">" + map.get("content") +"</p>"+
//                    "<img src='"+upload_mail_path +map.get("role")+"\\\\"+file.getOriginalFilename()+"' width = 180% height = 100%>"+
                    "<img src='"+upload_mail_path +map.get("role")+"\\\\"+file.getOriginalFilename()+"'>"+
                    "<p>以上，祝好！</p> " +
                    "<p>------------------------</p>" +
                    "<p>运营部_计划分部</p> " +
                    "<img src='"+upload_mail_path+"tianma_logo.png'>" +
                    " <p>厦门天马显示科技有限公司</p>" +
                    " <p>Xiamen Tianma Display Technology Co.,Ltd. </p>" +
                    "<p>福建省厦门市翔安区翔安西路6999号 No.6999,West Xiangan Road,XianganDistrict,Xiamen,China</p>" +
                    "</div></body></html>";
            Date day=new Date();
            SimpleDateFormat tmp_df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
            String imgUrl = "http://"+myhost+":8081/api/mail//"+map.get("role")+"/" + file.getOriginalFilename();
            String body_save = "<html>  <head>\n" +
                    "                <meta http-equiv=\"Content-Type\" content=\"text/html; charset=gbk\" />\n" +
                    "            </head><body><div>"+
                    "<p>" + map.get("content") +"</p>"+
//                    "<img src='"+upload_mail_path +map.get("role")+"\\\\"+file.getOriginalFilename()+"' width = 180% height = 100%>"+
                    "<img src='"+imgUrl+"'>"+
                    "<p><b >发布时间为："+ tmp_df.format(day)+"</b></p>"+
                    "</div></body></html>";
            FileUtil.write(upload_mail_path+map.get("role")+"\\\\public.html",body_save);
            
            Dispatch.put(mailItem, "HTMLBody", body);

            //附件
            Dispatch attachments = Dispatch.call(mailItem, "Attachments").getDispatch();
//            String filePath = upload_excel_path+map.get("role")+"\\\\"+map.get("role")+"_"+df.format(new Date())+".xlsx";
            String filePath = cur_excel_path;
//            if(map.get("role").equals("EVEN")){
                while(!isExist && time < 60){
                    cur_tmp = StringUtils.timeStamp2Date(tmp);
                    System.out.println(cur_tmp);
                    filePath_tmp = upload_excel_path+map.get("role")+"\\\\"+map.get("role")+"生产计划_"+cur_tmp.replace(".","")+".xlsx";
                    System.out.println("--------------------");
                    System.out.println(filePath_tmp);
                    file2 = new File(filePath_tmp);
                    if(!file2.exists()){
                        tmp = cur_time - 86400000;
                        cur_time = tmp;
                        time += 1;
                    }
                    else {
                        isExist = true;
                    }
                    cur_excel_path = filePath_tmp;
                    filePath = cur_excel_path;
                    System.out.println(filePath_tmp);
                }

            View_Mail view_mail = new View_Mail();
            view_mail.setTitle(map.get("title"));
            view_mail.setFilename(map.get("role")+"生产计划_"+cur_tmp.replace(".","")+".xlsx");
            if (uploadService.getMailByname(view_mail)) {
                uploadService.deleteMailInfo(view_mail);
                uploadService.insertMailInfo(map.get("title"),map.get("content"),format.format(cur_time2),map.get("role"),map.get("role")+"生产计划_"+cur_tmp.replace(".","")+".xlsx");
            } else {
                uploadService.insertMailInfo(map.get("title"),map.get("content"),format.format(cur_time2),map.get("role"),map.get("role")+"生产计划_"+cur_tmp.replace(".","")+".xlsx");
            }
//                uploadService.insertMailInfo(map.get("title"),map.get("content"),format.format(cur_time2),map.get("role"),map.get("role")+"生产计划_"+cur_tmp.replace(".","")+".xlsx");
//            }
            Dispatch.call(attachments, "Add" , filePath);
            Dispatch.call(mailItem, "Display");
            Dispatch.call(mailItem, "Send");
            System.out.println("1111111111111111111111111111111111111111111111111111111111111");
            return new Result("200", "success");
        }
        catch (Exception e) {
            System.out.println(("调用outlook失败,无法发送邮件"));
            System.out.println(e);
            return new Result("405", "fail");
        }

    }

    @RequestMapping("/getCurFile")
    public String getCurFile(String role){
        long cur_time = new Date().getTime();
        String cur_tmp = "";
        String filePath_tmp = "";
        Boolean isExist = false;
        String res = "";
        File file = new File(upload_excel_path);
        int time = 0;
        long tmp = cur_time;
        while(!isExist && time < 60){
            cur_tmp = StringUtils.timeStamp2Date(tmp);
            System.out.println(cur_tmp);
            res = role+"生产计划_"+cur_tmp.replace(".","")+".xlsx";
            filePath_tmp = upload_excel_path+role+"\\\\"+role+"生产计划_"+cur_tmp.replace(".","")+".xlsx";
            System.out.println("--------------------");
            System.out.println(filePath_tmp);
            file = new File(filePath_tmp);
            if(!file.exists()){
                tmp = cur_time - 86400000;
                cur_time = tmp;
                time += 1;
            }
            else {
                isExist = true;
            }

        }
        System.out.println("1111111111111");
        System.out.println(res);
        return res;
    }

    @RequestMapping(value = "/to_upload",method = RequestMethod.POST)
    @ResponseBody
    File download_up (MultipartFile file){
        if (file != null) {
            try {

                String filePath = upload_excel_path +"test"+"\\\\"+file.getOriginalFilename();
                FileInfo tmp = new FileInfo();
                tmp.setFilename(file.getOriginalFilename());
                System.out.println(filePath);
                File savedFile = new File(filePath);
                boolean isCreateSuccess = savedFile.createNewFile();
                // 是否创建文件成功
                if (isCreateSuccess) {
                    //将文件写入
                    file.transferTo(savedFile);
                    return savedFile;
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
        } else {
            System.out.println("文件是空的");
        }
        return null;
    }
    @RequestMapping(value = "/to_download", method = RequestMethod.GET)
    public void downloadExcel(HttpServletResponse response, String fileName) throws IOException {

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
            File f = new File(upload_excel_path+"test"+"\\\\"+fileName);

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


}
