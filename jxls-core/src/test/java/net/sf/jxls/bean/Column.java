package net.sf.jxls.bean;

/**
 * @author Leonid Vysochyn
 */
public class Column {
    String text;

    public Column(String text) {
        this.text = text;
    }

    public String getText() {
        return text;
    }

    public void setText(String text) {
        this.text = text;
    }
}
