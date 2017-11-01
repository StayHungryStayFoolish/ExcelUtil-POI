package com.excel.bean;

/**
 * Created by bonismo@hotmail.com
 * 上午10:25 on 17/11/1.
 * <p>
 * 如果生成 Excel 使用该类，根据 Excel 列建立对应的域
 */
public class ExportExcelBean {

    private Integer orgCode;

    private String orgName;

    private String userCode;

    private String userName;

    /**************** get/set 方法 **********************/
    public Integer getOrgCode() {
        return orgCode;
    }

    public void setOrgCode(Integer orgCode) {
        this.orgCode = orgCode;
    }

    public String getOrgName() {
        return orgName;
    }

    public void setOrgName(String orgName) {
        this.orgName = orgName;
    }

    public String getUserCode() {
        return userCode;
    }

    public void setUserCode(String userCode) {
        this.userCode = userCode;
    }

    public String getUserName() {
        return userName;
    }

    public void setUserName(String userName) {
        this.userName = userName;
    }
}
