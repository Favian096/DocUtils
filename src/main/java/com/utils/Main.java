package com.utils;

public class Main {

    //    设置测试文件基本路径
    public static final String PATH = System.getProperty("user.dir") +
            "\\src\\main\\resources\\";

    /*POM 右键settings中添加:
    <mirror>
        <id>com.e-iceblue</id>
        <url>http://repo.e-iceblue.cn/repository/maven-public/</url>
        <mirrorOf>com.e-iceblue</mirrorOf>
        <name>com.e-iceblue</name>
    </mirror>
    * */

    public static void main(String[] args) {
        System.out.println("Hello world!");
    }
}