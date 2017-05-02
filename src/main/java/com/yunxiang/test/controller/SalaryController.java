package com.yunxiang.test.controller;

import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;
import java.util.logging.Logger;

/**
 * Created by wangqingxiang on 2017/5/2.
 */
@Controller
public class SalaryController {

    @RequestMapping(value="/upload", method= RequestMethod.POST)
    public  String handleFileUpload(Model model,String  passWord,String fromEmail,@RequestParam("emails") MultipartFile emails,
                                                 @RequestParam("salarys") MultipartFile salarys){
        if(emails==null){
            return "请选择邮箱文件";
        }
        if(salarys==null){
            return "请选择工资文件";
        }

        Map<String, List<String>>  salaryMap= ReadExcelUtil.readSalary(salarys);
        Map<String, String> emailMap=ReadExcelUtil.readEmail(emails);

        System.out.println("开始读取excel文件");
        //读取excel

        if (salaryMap == null) {
            System.out.println("读取excel失败");
            return "error";
        }
        List<List<String>> results=new ArrayList<List<String>>();
        for (String userName : salaryMap.keySet()) {
            List<String> salary = salaryMap.get(userName);
            String email = emailMap.get(userName);
            if (email == null) {
                continue;
            }
            Map<String, Object> map = new HashMap<String, Object>();
            map.put("username", userName);
            map.put("salary", salary);
            map.put("date", new Date());
            map.put("fromEmail",fromEmail);
            map.put("fromName","孙冰冰");
            map.put("passWord",passWord);
            String templatePath = "salary.ftl";
            boolean result=SendMailUtil.sendFtlMail( email, "新兴投资有限公司工资明细", templatePath, map);

            if(!result){
                salary=new ArrayList<String>();
                salary.add(userName);
                salary.add("失败");
            }
            results.add(salary);

        }
        model.addAttribute("salarys",results);
        return "email";
    }
}
