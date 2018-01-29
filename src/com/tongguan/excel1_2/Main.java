package com.tongguan.excel1_2;

import java.util.Date;

public class Main {

    public static void main(String[] args) {
        System.out.println("Hello World!");
        long starTime = new Date().getTime();

        AccountAge accountAge = new AccountAge("三年其他应收账款");
        accountAge.doInit();

        long endTime = new Date().getTime();
        System.out.println("花费时间"+(endTime-starTime));
    }
}
