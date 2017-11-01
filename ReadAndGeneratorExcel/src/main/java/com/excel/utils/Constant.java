package com.excel.utils;

/**
 * Created by bonismo@hotmail.com
 * 上午11:07 on 17/11/1.
 * <p>
 * 存放读取、写入 Excel 列对应 Bean 的域
 * 目前一样，根据实际需求，更改
 */
public interface Constant {

    /************** 读取 Excel 对应 Bean 的域 **********************/
    String IMPORT_ORG_CODE = "orgCode";
    String IMPORT_ORG_NAME = "orgName";
    String IMPORT_USER_CODE = "userCode";
    String IMPORT_USER_NAME = "userName";

    /*************** 写入 Excel 对应 Bean 的域 *********************/
    String EXPORT_ORG_CODE = "orgCode";
    String EXPORT_ORG_NAME = "orgName";
    String EXPORT_USER_CODE = "userCode";
    String EXPORT_USER_NAME = "userName";
}
