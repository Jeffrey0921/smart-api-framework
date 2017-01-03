package com.qa.framework.exception;

/**
 * Created by Administrator on 2017/1/3.
 */
public class NoSuchNameInExcelException extends Exception {
    public NoSuchNameInExcelException(String name) {
        super("在Excel表格中找不到" + name + "这个对象所对应的名称");
    }
}
