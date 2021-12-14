package com.example.demo.service.spi;

public class Chinese implements Person {

    @Override
    public String sayHello() {
        System.out.println("China");
        return "China";
    }
}
