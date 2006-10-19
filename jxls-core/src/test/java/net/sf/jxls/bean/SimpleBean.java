package net.sf.jxls.bean;

import java.util.Date;

/**
 * @author Leonid Vysochyn
 */
public class SimpleBean {
    private String name;
    private Double doubleValue;
    private Integer intValue;
    private Date dateValue;
    private SimpleBean other;



    public SimpleBean(String name, Double doubleValue, Integer intValue, Date dateValue) {
        this.name = name;
        this.doubleValue = doubleValue;
        this.intValue = intValue;
        this.dateValue = dateValue;
    }

    public SimpleBean(String name, Double doubleValue, Integer intValue) {
        this.name = name;
        this.doubleValue = doubleValue;
        this.intValue = intValue;
    }

    public String getBeansProp(){
        return "beans_for_" + name;
    }


    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Double getDoubleValue() {
        return doubleValue;
    }

    public void setDoubleValue(Double doubleValue) {
        this.doubleValue = doubleValue;
    }

    public SimpleBean getOther() {
        return other;
    }

    public void setOther(SimpleBean other) {
        this.other = other;
    }

    public Integer getIntValue() {
        return intValue;
    }

    public void setIntValue(Integer intValue) {
        this.intValue = intValue;
    }

    public Date getDateValue() {
        return dateValue;
    }

    public void setDateValue(Date dateValue) {
        this.dateValue = dateValue;
    }
}
