package net.sf.jxls.sample;


import net.sf.jxls.processor.PropertyPreprocessor;

import java.util.HashSet;
import java.util.Set;

/**
 * @author Leonid Vysochin
 */
public class EmptyCellPreprocessor implements PropertyPreprocessor {

    private Set hiddenProperties = new HashSet();

    public boolean addHiddenProperty(String propertyTemplateName) {
        return hiddenProperties.add(propertyTemplateName);
    }

    /**
     * if the property is private we return empty string to indicate it should not be visible
     *
     * @param propertyTemplateName
     * @return replacement value for given property
     */
    public String processProperty(String propertyTemplateName) {
        if (hiddenProperties.contains(propertyTemplateName)) {
            return "";
        } else {
            return null;
        }
    }
}
