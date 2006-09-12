package net.sf.jxls.parser;

import org.apache.commons.jexl.ExpressionFactory;
import org.apache.commons.jexl.JexlContext;
import org.apache.commons.jexl.JexlHelper;
import net.sf.jxls.parser.Property;
import net.sf.jxls.transformer.Configuration;

import java.util.Map;
import java.util.List;
import java.util.ArrayList;


/**
 * Represents JEXL expression
 * @author Leonid Vysochyn
 */
public class Expression {
    String expression;
    Map beans;
    List properties = new ArrayList();
    org.apache.commons.jexl.Expression jexlExpresssion;

    Configuration config;

    Property collectionProperty;

    public Property getCollectionProperty() {
        return collectionProperty;
    }

    public List getProperties() {
        return properties;
    }

    public String getExpression() {
        return expression;
        
    }

    public Expression(String expression, Configuration config) {
        this.expression = expression;
        this.config = config;
    }

    public Expression(String expression, Map beans, Configuration config) throws Exception {
        this.config = config;
        this.expression = expression;
        this.beans = beans;
        jexlExpresssion = ExpressionFactory.createExpression( expression );
        parse();
    }

    public Object evaluate() throws Exception {
          
        if (beans != null && !beans.isEmpty()){
            JexlContext context = JexlHelper.createContext();
            context.setVars(beans);
            return jexlExpresssion.evaluate(context);
        } else {
            return expression;
        }
    }

    private void parse() {
        Property prop = new Property(expression, beans, config);
        this.properties = new ArrayList();
        this.properties.add(prop);
        if (prop.isCollection() && collectionProperty == null) {
            this.collectionProperty = prop;
        }
    }

    public String toString() {
        return expression;
    }
}