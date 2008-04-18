package net.sf.jxls.bean;

import java.util.ArrayList;
import java.util.List;

/**
 * @author Leonid Vysochyn
 */
public class BeanWithList {

    private String name;
    private Double doubleValue;

    private List beans = new ArrayList();

    public BeanWithList(String name, Double doubleValue) {
        this.name = name;
        this.doubleValue = doubleValue;
    }

    public BeanWithList(String name) {
        this.name = name;
    }

    public void addBean(SimpleBean bean){
        beans.add( bean );
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

    public List getBeans() {
        return beans;
    }

    public void setBeans(List beans) {
        this.beans = beans;
    }
}
