package com.example.demo.service.spi;

import java.util.ServiceLoader;

public class SpiTest {
    public static void main(String[] args) {
        //java-spi参考 https://www.jianshu.com/p/46b42f7f593c
        ServiceLoader<Person> personServiceLoader = ServiceLoader.load(Person.class);
        for (Person p : personServiceLoader) {
            p.sayHello();
        }
    }
}
