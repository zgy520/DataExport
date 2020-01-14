package com.excel.tool.core.model;

import lombok.Data;
import lombok.extern.slf4j.Slf4j;

import java.util.Date;

@Data
@Slf4j
public class User {
    private String name;
    private String birthDate;
    private int age;
    private String address;

    public User(String name,String birthDate,int age,String address){
        this.name = name;
        this.birthDate = birthDate;
        this.age = age;
        this.address = address;
    }
}
