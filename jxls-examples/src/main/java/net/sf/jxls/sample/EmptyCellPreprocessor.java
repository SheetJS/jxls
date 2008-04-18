package net.sf.jxls.sample;


import java.util.HashSet;
import java.util.Set;

import net.sf.jxls.processor.PropertyPreprocessor;

/**
 * @author Leonid Vysochin
 */
public class EmptyCellPreprocessor implements PropertyPreprocessor {

    private Set hiddenProperties = new HashSet();

    public boolean addHiddenProperty(String propertyTemplateName) {
        return hiddenProperties.add(propertyTemplateName);
    }

    /**
     * If the property is private we return empty string to indicate it should not be visible
     * @param propertyTemplateName - The name of the property as it is in template file
     * @return replacement value for given property
     */
    public String processProperty(String propertyTemplateName) {
        if (hiddenProperties.contains(propertyTemplateName)) {
            return "";
        }
        return null;
    }
}
