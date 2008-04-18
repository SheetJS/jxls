package net.sf.jxls.bean;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * @author Leonid Vysochyn
 */
public class MyBean {
    private int id;
    private String name = "My Bean Name";
    private boolean flag = true;
    private List collection = new ArrayList();
    private String[] myArray = new String[2];
    private Date date = new Date();

    private List innerCollection = new ArrayList();

    public MyBean() {
        myArray[0] = "first";
        myArray[1] = "last";
    }

    public MyBean(String name) {
        this();
        this.name = name;
    }

    public MyBean(int id) {
        this();
        this.id = id;
    }

    public void addBean(MyBean bean){
        collection.add( bean );
    }

    public int getId() {
        return id;
    }

    public String getName() {
        return name;
    }

    public boolean getFlag() {
        return flag;
    }

    public Date getDate() {
        return date;
    }

    public String[] getMyArray() {
        return myArray;
    }

    public List getCollection() {
        return collection;
    }

    public List getInnerCollection() {
        return innerCollection;
    }


    public String printIt() {
        return myArray[0] + " is the first array element.";
    }

    public String echo(String str) {
        return str;
    }

}
