package net.sf.jxls.parser;

import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.Map;

import net.sf.jxls.transformer.Configuration;

import org.apache.commons.jexl.JexlContext;
import org.apache.commons.jexl.JexlHelper;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;

/**
 * Represents a property in excel template
 * @author Leonid Vysochyn
 */
public class Property {
    protected final Log log = LogFactory.getLog(getClass());
    List propertyTokens = new ArrayList();

    private String beanName;
    private Object bean;
    private String collectionName;
    private Collection collection;

    private String property;
    private Object propertyValue;

    Configuration config;

    public Property(String value) {
        propertyValue = value;
    }

    public Property(String property, Map beans, Configuration config) {
        this.property = property;
        this.config = config;
        propertyValue = getPropertyValue(beans);
    }

    public boolean isConstant(){
        return property==null;
    }

    public Object getPropertyValue(Map beans) {
        
        JexlContext context = JexlHelper.createContext();
        context.setVars(beans);
        ExpressionCollectionParser parser = new ExpressionCollectionParser(context, this.property + ";", config.isJexlInnerCollectionsAccess());
        if (parser.getCollection() == null) {
            propertyValue = null; 
        } else {
            collectionName = parser.getCollectionExpression();
            collection = parser.getCollection();
            beanName = null;
            bean = null;
        }
        
        return propertyValue;
    }

    public boolean isCollection() {
        return collectionName != null;
    }

    public boolean isNull() {
        return getPropertyValue() == null;
    }

    public String getBeanName() {
        return beanName;
    }

    public String getCollectionName() {
        return collectionName;
    }

    public void setCollectionName(String collectionName) {
        this.collectionName = collectionName;
    }


    public String getProperty() {
        return property;
    }


    public Collection getCollection() {
        return collection;
    }

    public void setCollection(Collection collection) {
        this.collection = collection;
    }

    public String getFullCollectionName() {
        if (beanName == null) {
            return collectionName;
        } else {
            return beanName + "." + collectionName;
        }
    }

    public String getPropertyNameAfterLastDot() {
        String propertyName = null;
        if (property != null) {
            int dotIndex = property.lastIndexOf(".");
            if (dotIndex >= 0) {
                propertyName = property.substring(dotIndex + 1);
            } else {
                propertyName = property;
            }
        }
        return propertyName;
    }

    public String getPropertyNameAfterFirstDot() {
        String propertyName = null;
        if (property != null) {
            int dotIndex = property.indexOf(".");
            if (dotIndex >= 0) {
                propertyName = property.substring(dotIndex + 1);
            } else {
                propertyName = property;
            }
        }
        return propertyName;
    }

    public String toString() {
        return "Property{" +
                "property='" + property + "'}";
    }

    public Object getPropertyValue() {
        if( bean instanceof String){
            return (String)bean;
        }else{
            return propertyValue;
        }
    }

    public void setPropertyValue(Object propertyValue) {
        this.propertyValue = propertyValue;
    }

    public Object getBean() {
        return bean;
    }

    public void setBean(Object bean) {
        this.bean = bean;
    }

}
