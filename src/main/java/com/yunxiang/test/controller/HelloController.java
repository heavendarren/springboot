package com.yunxiang.test.controller;

import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;

import java.util.HashMap;
import java.util.Map;

/**
 * Created by wangqingxiang on 2017/2/17.
 */
@Controller
public class HelloController {
    @RequestMapping("/index")
    public String index(Model model) {
        model.addAttribute("name", "小冰冰");
        return "index";
    }

    @RequestMapping("/json")
    @ResponseBody
    public Map<String, Object> json() {
        Map<String, Object> map = new HashMap<String, Object>();
        map.put("name", "yunxiang");
        map.put("sex", "man");
        map.put("age", "18");
        return map;
    }
}
