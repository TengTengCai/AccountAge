package com.tongguan.excel1_2;

/**
 *  dataBean实例对象
 */
public class DataBean {
    private String customer;    //客户名称
    private String titleName;   //标题名称
    private String year;        //年份
    private Double value;       //值
    private String error;       //错误信息
    private boolean isError = false;    //是否有误

    /**
     * 构造方法
     * @param customer  客户名称
     * @param titleName 标题
     * @param year      年份
     * @param value     值
     */
    DataBean(String customer, String titleName, String year, double value) {
        this.customer = customer;
        this.titleName = titleName;
        this.year = year;
        this.value = value;
    }

    /**
     * 单元格错误的构造方法
     * @param customer  客户名称
     * @param titleName 标题
     * @param year      年份
     * @param error     错误信息
     */
    DataBean(String customer, String titleName, String year, String error) {
        this.customer = customer;
        this.titleName = titleName;
        this.year = year;
        this.error = error;
        isError = true;
    }

    public boolean isError() {
        return isError;
    }

    public String getError() {
        return error;
    }

    public void setError(String error) {
        this.error = error;
    }

    public DataBean() {
    }

    public String getCustomer() {
        return customer;
    }

    public String getTitleName() {
        return titleName;
    }

    public String getYear() {
        return year;
    }

    public Double getValue() {
        return value;
    }

    public void setCustomer(String customer) {
        this.customer = customer;
    }

    public void setTitleName(String titleName) {
        this.titleName = titleName;
    }

    public void setYear(String year) {
        this.year = year;
    }

    public void setValue(Double value) {
        this.value = value;
    }
}
