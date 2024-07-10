package com.mysite.spring_erp.dashboard.controller;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestParam;

@RequestMapping("/dashboard")
@Controller
public class DashController {
    @GetMapping("/index")
    public String dash_board() {
        return "dashboard/index";
    }

}
