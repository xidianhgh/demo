package com.example.demo.service.spi;

public class American implements Person {
    @Override
    public String sayHello() {
        System.out.println("USA");
        return "USA";
    }
}
