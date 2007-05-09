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
}