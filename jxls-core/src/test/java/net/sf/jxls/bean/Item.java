package net.sf.jxls.bean;

import java.util.ArrayList;
import java.util.List;

/**
 * @author Leonid Vysochyn
 */
public class Item {
    private List attributes = new ArrayList();

    public Item(String name) {
        attributes.add(name + " Attribute 1");
        attributes.add(name + " Attribute 2");
    }

    //getters and setters
    public List getAttributes() {
        return attributes;
    }

    public void setAttributes(List attributes) {
        this.attributes = attributes;
    }

    private String key;
    private List values = new ArrayList();
    public Item(String key, int[] _values) {
        this.key = key;
        for (int i = 0; i < _values.length; i++) {
            values.add(Integer.valueOf(_values[i]));
        }
    }
    public String getKey() { return key; }
    public List getValues() { return values; }    

}