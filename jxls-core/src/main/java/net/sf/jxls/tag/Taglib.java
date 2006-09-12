package net.sf.jxls.tag;

import java.util.Map;
import java.util.HashMap;

/**
 * Defines mapping between java class files and jx tag names
 * @author Leonid Vysochyn
 */
public class Taglib {
    static String[] tagName = new String[]{ "forEach", "if", "outline"};
    static String[] tagClassName = new String[]{ "net.sf.jxls.tag.ForEachTag", "net.sf.jxls.tag.IfTag", "net.sf.jxls.tag.OutlineTag" };

    static Map tags = new HashMap();

    static{
        for (int i = 0; i < tagName.length; i++) {
            String key = tagName[i];
            String value = tagClassName[i];
            tags.put( key, value );
        }
    }

    static public Map getTags(){
        return tags;
    }


}
