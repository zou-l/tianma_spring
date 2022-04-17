package com.tianma.yunying.util;
//import com.tianma.yunying.entity.Email;
//import com.xiaofei.emaildemo.controller.MyAuthenticator;
//import com.xiaofei.emaildemo.pojo.Email;
import org.springframework.beans.factory.annotation.Value;
import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.*;
import javax.mail.internet.*;
import java.io.File;
import java.io.UnsupportedEncodingException;
import java.util.Date;
import java.util.Properties;

public class EmailUtil {
//    /**
//     * 以HTML格式发送邮件
//     * @param mailInfo 待发送的邮件信息
//     */
//    public static boolean sendHtmlMail(Email mailInfo){
//        // 判断是否需要身份认证
//        MyAuthenticator authenticator = null;
//        Properties pro = mailInfo.getProperties();
//        //如果需要身份认证，则创建一个密码验证器
//        if (mailInfo.isValidate()) {
//            authenticator = new MyAuthenticator(mailInfo.getUserName(), mailInfo.getPassword());
//        }
//        // 根据邮件会话属性和密码验证器构造一个发送邮件的session
//        Session sendMailSession = Session.getDefaultInstance(pro);
//        try {
//            // 根据session创建一个邮件消息
//            Message mailMessage = new MimeMessage(sendMailSession);
//            // 创建邮件发送者地址
//            Address from = new InternetAddress(mailInfo.getFromAddress());
//            // 设置邮件消息的发送者
//            mailMessage.setFrom(from);
//            // 创建邮件的接收者地址，并设置到邮件消息中
//            Address to = new InternetAddress(mailInfo.getToAddress());
//            // Message.RecipientType.TO属性表示接收者的类型为TO
//            mailMessage.setRecipient(Message.RecipientType.TO,to);
//            // 设置邮件消息的主题
//            mailMessage.setSubject(mailInfo.getSubject());
//            // 设置邮件消息发送的时间
//            mailMessage.setSentDate(new Date());
//            // MiniMultipart类是一个容器类，包含MimeBodyPart类型的对象
//            Multipart mainPart = new MimeMultipart();
//
//            // 创建一个包含HTML内容的MimeBodyPart
//            BodyPart html = new MimeBodyPart();
//            // 设置HTML内容
//            html.setContent(mailInfo.getContent(), "text/html; charset=utf-8");
//
//            // 创建一个包含文件内容的MimeBodyPart
//            BodyPart file  = new MimeBodyPart();
//            DataHandler handler = new DataHandler(new FileDataSource(new File("/ui/" + mailInfo.getAttachFileNames())));
//            file.setDataHandler(handler);
//            file.setFileName(mailInfo.getAttachFileNames());
//
//            //拼装邮件正文
//            mainPart.addBodyPart(html);
//            mainPart.addBodyPart(file);
//
//            // 将MiniMultipart对象设置为邮件内容
//            mailMessage.setContent(mainPart);
//            // 发送邮件
//            Transport.send(mailMessage);
//            return true;
//        } catch (MessagingException ex) {
//            ex.printStackTrace();
//        }
//        return false;
//    }
    public static MimeMessage createMimeMessage(Session session, String sendMail, String receiveMail,String subject, String Content,String attachmentPath) throws Exception {
        // 1. 创建一封邮件
        MimeMessage message = new MimeMessage(session);

        // 2. From: 发件人
        message.setFrom(new InternetAddress(sendMail, "计划系统", "UTF-8"));

        // 3. To: 收件人（可以增加多个收件人、抄送、密送）
        message.setRecipient(MimeMessage.RecipientType.TO, new InternetAddress(receiveMail, "XX用户", "UTF-8"));
//        message.setRecipient(MimeMessage.RecipientType.TO, new InternetAddress(receiveMail, "XX用户2", "UTF-8"));

        // 4. Subject: 邮件主题
        message.setSubject(subject, "UTF-8");

        // 5. Content: 邮件正文（可以使用html标签）
        message.setContent(Content, "text/html;charset=UTF-8");
        // 6. 设置发件时间
        message.setSentDate(new Date());

        // 7. 保存设置
//        message.saveChanges();

//        Multipart multipart = new MimeMultipart();
//        BodyPart messageBodyPart = new MimeBodyPart();
////        messageBodyPart = new MimeBodyPart();
//        DataSource source = new FileDataSource(attachmentPath);
//        messageBodyPart.setDataHandler(new DataHandler(source));
//        String attachmentPathTrim = attachmentPath.trim();
//        String fileName = attachmentPathTrim.substring(attachmentPathTrim.lastIndexOf("\\")+1);
//        messageBodyPart.setFileName(fileName);
//        System.out.println(fileName);
//        multipart.addBodyPart(messageBodyPart);
//        message.setSentDate(new Date());
//        message.saveChanges();
        MimeMultipart multipart = new MimeMultipart();
        //读取本地图片,将图片数据添加到"节点"
        MimeBodyPart image = new MimeBodyPart();
        DataHandler dataHandler1 = new DataHandler(new FileDataSource("D:\\CodePath\\test\\testpic.png"));
        image.setDataHandler(dataHandler1);
        image.setContentID("image_suo");
        //创建文本节点
        MimeBodyPart text = new MimeBodyPart();
        text.setContent(Content+"这张图片是<br/><img src='cid:image_suo'/>","text/html;charset=UTF-8");

        //将文本和图片添加到multipart
        multipart.addBodyPart(text);
        multipart.addBodyPart(image);

//        multipart.setSubType("related");//关联关系

        MimeBodyPart file1 = new MimeBodyPart();
        DataHandler dataHandler2 = new DataHandler(new FileDataSource(attachmentPath));
        file1.setDataHandler(dataHandler2);
        multipart.addBodyPart(file1);
        multipart.setSubType("mixed");//混合关系
        file1.setFileName(MimeUtility.encodeText(dataHandler2.getName()));
        message.setContent(multipart);

        message.setSentDate(new Date());
        message.saveChanges();

        return message;
    }
}
