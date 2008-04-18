package net.sf.jxls.bean;

import java.util.ArrayList;
import java.util.List;

/**
 * @author Leonid Vysochyn
 */
public class Bean {
    private String name;
    private List collection = new ArrayList();

    public Bean() {
        name = "test";
        InnerBean bean1 = new InnerBean("inner1");
        bean1.getInnerCollection().add("1");
        bean1.getInnerCollection().add("2");
        bean1.getInnerCollection().add("3");
        bean1.getInnerCollection().add("4");
        collection.add(bean1);
        InnerBean bean2 = new InnerBean("inner2");
        bean2.getInnerCollection().add("a");
        bean2.getInnerCollection().add("b");
        bean2.getInnerCollection().add("c");
        bean2.getInnerCollection().add("d");
        collection.add(bean2);
        InnerBean bean3 = new InnerBean("inner3");
        bean3.getInnerCollection().add("i");
        bean3.getInnerCollection().add("ii");
        bean3.getInnerCollection().add("iii");
        bean3.getInnerCollection().add("iv");
        collection.add(bean3);
    }

    public String getName() {
        return name;
    }

    public List getCollection() {
        return collection;
    }

    public static class InnerBean {
        String name;
        List innerCollection = new ArrayList();

        public InnerBean(String name) {
            this.name = name;
        }

        public String getName() {
            return name;
        }

        public List getInnerCollection() {
            return innerCollection;
        }
    }
}
