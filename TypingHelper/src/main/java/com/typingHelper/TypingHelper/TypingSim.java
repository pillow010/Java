package com.typingHelper.TypingHelper;

import java.awt.*;
import java.awt.event.KeyEvent;

public class TypingSim {
    public static void main(String[] args) {
        try {
            Robot robot = new Robot();
            robot.keyPress(KeyEvent.VK_ALT);
            robot.keyPress(KeyEvent.VK_TAB);
            robot.keyRelease(KeyEvent.VK_TAB);
            robot.keyRelease(KeyEvent.VK_ALT);
        } catch (AWTException e) {
            e.printStackTrace();
        }
    }
}
