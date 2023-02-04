package com.typingHelper.TypingHelper;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;

@Controller
public class MyController {
    @RequestMapping("/")
    public String index() {
        return "MainPage";
    }

    @ModelAttribute("inputText")
    public String inputText() {
        return "";
    }

    @PostMapping("/submitForm")
    public String submitForm(@ModelAttribute("inputText") String inputText) {
        // Do something with the input text
        return "redirect:/";
    }
}

